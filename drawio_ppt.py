"""
drawio_ppt.py — Bidirectional converter between Draw.io XML and PowerPoint.

Usage:
    python drawio_ppt.py drawio2ppt  input.drawio  output.pptx  [--config config.ini]
    python drawio_ppt.py ppt2drawio  input.pptx    output.drawio

drawio2ppt: Parses .drawio XML (or raw mxGraphModel XML), creates a .pptx
            with native shapes, connectors, colors, and labels.  Optionally
            snaps colours to a corporate palette via --config.

ppt2drawio: Reads a .pptx, extracts every shape/connector on every slide,
            and writes a .drawio file (one diagram page per slide).
"""

from __future__ import annotations

import argparse
import base64
import configparser
import math
import re
import sys
import zlib
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.util import Emu, Pt


# ---------------------------------------------------------------------------
# Drawio encoding helpers (ported from drawio_tools.py)
# ---------------------------------------------------------------------------

def decode_diagram_data(data: str) -> str:
    raw = base64.b64decode(data)
    raw = zlib.decompressobj(-15).decompress(raw)
    return unquote(raw.decode())


def encode_diagram_data(xml: str) -> str:
    from urllib.parse import quote
    encoded = quote(xml, safe="~()*!.'")
    compressed = zlib.compressobj(
        zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15,
        memLevel=8, strategy=zlib.Z_DEFAULT_STRATEGY,
    ).compress(encoded.encode())
    compressed += zlib.compressobj(
        zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15,
        memLevel=8, strategy=zlib.Z_DEFAULT_STRATEGY,
    ).flush()
    # simpler: do it in one shot
    c = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15)
    out = c.compress(encoded.encode()) + c.flush()
    return base64.b64encode(out).decode()


# ---------------------------------------------------------------------------
# HTML-to-text helper
# ---------------------------------------------------------------------------

class _StripHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag in ("br", "div", "p"):
            self._parts.append("\n")

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self) -> str:
        # Collapse multiple newlines, strip leading/trailing
        import re
        text = "".join(self._parts)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


def strip_html(html: str) -> str:
    p = _StripHTML()
    p.feed(html)
    return p.get_text()


# ---------------------------------------------------------------------------
# Style parsing
# ---------------------------------------------------------------------------

def parse_style(style_str: str) -> dict[str, str]:
    """Parse a drawio style string like 'rounded=1;fillColor=#dae8fc;...' into a dict."""
    result: dict[str, str] = {}
    if not style_str:
        return result
    for token in style_str.rstrip(";").split(";"):
        if "=" in token:
            k, v = token.split("=", 1)
            result[k.strip()] = v.strip()
        elif token.strip():
            result[token.strip()] = ""
    return result


def build_style(props: dict[str, str]) -> str:
    parts = []
    for k, v in props.items():
        parts.append(f"{k}={v}" if v else k)
    return ";".join(parts) + ";"


# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------

def hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    if len(h) < 6:
        return (0, 0, 0)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"#{r:02X}{g:02X}{b:02X}"


def colour_distance(c1: tuple[int, int, int], c2: tuple[int, int, int]) -> float:
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(c1, c2)))


def load_palette(ini_path: str) -> list[str]:
    """Load colour hex codes from config.ini [colours] section."""
    cfg = configparser.ConfigParser()
    cfg.read(ini_path)
    colours = []
    if cfg.has_section("colours"):
        for key in cfg["colours"]:
            val = cfg["colours"][key]
            hex_code = val.split("|")[0].strip()
            colours.append(hex_code)
    return colours


def snap_colour(hex_colour: str, palette: list[str]) -> str:
    if not palette:
        return hex_colour
    target = hex_to_rgb(hex_colour)
    best = min(palette, key=lambda c: colour_distance(target, hex_to_rgb(c)))
    return best


def darken(hex_colour: str, amount: int = 20) -> str:
    r, g, b = hex_to_rgb(hex_colour)
    return rgb_to_hex(max(0, r - amount), max(0, g - amount), max(0, b - amount))


# ---------------------------------------------------------------------------
# Drawio XML parsing
# ---------------------------------------------------------------------------

def load_drawio_xml(path: str) -> ET.Element:
    """Load a .drawio or raw mxGraphModel XML file. Returns the mxGraphModel element."""
    tree = ET.parse(path)
    root = tree.getroot()

    # If it's a full .drawio (mxfile wrapper), decode diagram data
    if root.tag == "mxfile":
        diagram = root.find("diagram")
        if diagram is not None:
            text = (diagram.text or "").strip()
            if text:
                # Compressed/encoded diagram
                try:
                    xml_str = decode_diagram_data(text)
                    return ET.fromstring(xml_str)
                except Exception:
                    pass
            # Uncompressed — mxGraphModel is a child element
            mg = diagram.find("mxGraphModel")
            if mg is not None:
                return mg

    # Raw mxGraphModel XML (like taxonomy.drawio.xml)
    if root.tag == "mxGraphModel":
        return root

    raise ValueError(f"Cannot find mxGraphModel in {path}")


def get_absolute_coords(nodes_by_id: dict, node_el: ET.Element) -> tuple[float, float]:
    """Walk the parent chain to accumulate absolute x, y."""
    x, y = 0.0, 0.0
    current = node_el
    while current is not None:
        geo = current.find("mxGeometry")
        if geo is not None:
            x += float(geo.get("x", 0))
            y += float(geo.get("y", 0))
        pid = current.get("parent", "0")
        if pid == "0" or pid not in nodes_by_id:
            break
        current = nodes_by_id[pid]
    return x, y


def resolve_label(node_el: ET.Element, parent_el: Optional[ET.Element]) -> str:
    """Extract the human-readable label, resolving %placeholder% tokens."""
    parts = []

    # Label from <object> parent
    if parent_el is not None and parent_el.tag == "object":
        lbl = parent_el.get("label", "")
        if lbl:
            parts.append(strip_html(lbl))

    # Value from the mxCell itself
    val = node_el.get("value", "")
    if val:
        parts.append(strip_html(val))

    label = "\n".join(p for p in parts if p)

    # Resolve %placeholder% tokens from the <object> parent
    if "%" in label and parent_el is not None:
        def _replace(m):
            key = m.group(1)
            return parent_el.get(key, key)
        label = re.sub(r"%(\w+)%", _replace, label)

    # Append description if present
    desc_el = parent_el if parent_el is not None else node_el
    desc = desc_el.get("description", "")
    if desc:
        label = label + "\n" + desc.strip()

    return label


# ---------------------------------------------------------------------------
# drawio → pptx
# ---------------------------------------------------------------------------

# Drawio uses pixels; PowerPoint uses EMU (914400 per inch, 12700 per pt).
# Drawio's coordinate system is roughly 1 pixel = 1 CSS pixel ≈ 0.75pt.
DRAWIO_PX_TO_EMU = int(0.75 * 12700)  # 9525 EMU per drawio pixel


def dx(val: float) -> int:
    return int(val * DRAWIO_PX_TO_EMU)


def drawio_to_pptx(
    drawio_path: str,
    pptx_path: str,
    config_path: Optional[str] = None,
    exclude_layers: Optional[list[str]] = None,
    back_layers: Optional[list[str]] = None,
    line_darken: int = 20,
) -> None:
    palette = load_palette(config_path) if config_path else []
    exclude_layers = [l.lower() for l in (exclude_layers or [])]
    back_layers = [l.lower() for l in (back_layers or [])]

    mg = load_drawio_xml(drawio_path)
    root_el = mg.find("root")
    if root_el is None:
        raise ValueError("No <root> element in mxGraphModel")

    # Index all elements by id
    all_elements: dict[str, ET.Element] = {}
    for el in root_el.iter():
        eid = el.get("id")
        if eid:
            all_elements[eid] = el

    # Build parent map (child → parent element)
    parent_map: dict[str, ET.Element] = {}
    for el in root_el.iter():
        for child in el:
            cid = child.get("id")
            if cid:
                parent_map[cid] = el

    # Determine which IDs belong to excluded layers
    def is_excluded(el: ET.Element) -> bool:
        current = el
        while current is not None:
            val = (current.get("value") or "").strip().lower()
            if val and val in exclude_layers:
                return True
            pid = current.get("parent", "0")
            if pid == "0" or pid not in all_elements:
                break
            current = all_elements[pid]
        return False

    def is_back_layer(el: ET.Element) -> bool:
        current = el
        while current is not None:
            val = (current.get("value") or "").strip().lower()
            if val and val in back_layers:
                return True
            pid = current.get("parent", "0")
            if pid == "0" or pid not in all_elements:
                break
            current = all_elements[pid]
        return False

    # Collect vertices and edges, tracking which <object> wraps each mxCell.
    # mxCells inside <object> elements inherit the object's id and attributes.
    vertices: list[tuple[ET.Element, Optional[ET.Element]]] = []  # (mxCell, object_parent)
    edges: list[tuple[ET.Element, Optional[ET.Element]]] = []
    seen_cells: set[int] = set()  # avoid duplicates by element identity

    # First pass: mxCells inside <object> wrappers
    for obj in root_el.iter("object"):
        if is_excluded(obj):
            continue
        for cell in obj.findall("mxCell"):
            if id(cell) in seen_cells:
                continue
            seen_cells.add(id(cell))
            if cell.get("vertex") == "1":
                vertices.append((cell, obj))
            elif cell.get("edge") == "1":
                edges.append((cell, obj))

    # Second pass: standalone mxCells (not inside objects)
    for el in root_el.iter("mxCell"):
        if id(el) in seen_cells:
            continue
        seen_cells.add(id(el))
        if is_excluded(el):
            continue
        if el.get("vertex") == "1":
            vertices.append((el, None))
        elif el.get("edge") == "1":
            edges.append((el, None))

    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 16:9
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    shape_map: dict[str, object] = {}  # drawio id → pptx shape

    # --- bounding box pass ---
    bb_min_x, bb_min_y = float("inf"), float("inf")
    bb_max_x, bb_max_y = float("-inf"), float("-inf")
    for v, _obj in vertices:
        geo = v.find("mxGeometry")
        if geo is None:
            continue
        ax, ay = get_absolute_coords(all_elements, v)
        w = float(geo.get("width", 50))
        h = float(geo.get("height", 50))
        bb_min_x = min(bb_min_x, ax)
        bb_min_y = min(bb_min_y, ay)
        bb_max_x = max(bb_max_x, ax + w)
        bb_max_y = max(bb_max_y, ay + h)

    # NOTE: Don't include edge mxPoints in bounding box — they can be wildly
    # off-screen and distort the layout. Vertices define the real diagram extent.

    diag_w = bb_max_x - bb_min_x or 1
    diag_h = bb_max_y - bb_min_y or 1
    margin_emu = Pt(30)  # 30pt margin (~10mm)
    avail_w = prs.slide_width - 2 * margin_emu
    avail_h = prs.slide_height - 2 * margin_emu
    scale_x = avail_w / dx(diag_w)
    scale_y = avail_h / dx(diag_h)
    scale = min(scale_x, scale_y, 1.0)  # don't upscale beyond 1:1

    # Center the diagram on the slide
    rendered_w = int(dx(diag_w) * scale)
    rendered_h = int(dx(diag_h) * scale)
    offset_x = (prs.slide_width - rendered_w) // 2
    offset_y = (prs.slide_height - rendered_h) // 2

    def to_emu_x(val: float) -> int:
        return int((dx(val - bb_min_x)) * scale + offset_x)

    def to_emu_y(val: float) -> int:
        return int((dx(val - bb_min_y)) * scale + offset_y)

    def to_emu_w(val: float) -> int:
        return int(dx(val) * scale)

    # --- Sort vertices: largest shapes first (containers behind), smallest on top ---
    def _shape_area(item):
        v, _ = item
        geo = v.find("mxGeometry")
        if geo is None:
            return 0
        return float(geo.get("width", 50)) * float(geo.get("height", 50))

    vertices.sort(key=_shape_area, reverse=True)

    # --- Classify shapes ---
    MIN_SHAPE_AREA = 1500  # skip tiny annotation shapes (< ~39x39 drawio px)
    from pptx.enum.text import PP_ALIGN
    import re as _re

    rendered_count = 0
    skipped_count = 0

    for v, obj_parent in vertices:
        geo = v.find("mxGeometry")
        if geo is None:
            continue

        ax, ay = get_absolute_coords(all_elements, v)
        w = float(geo.get("width", 50))
        h = float(geo.get("height", 50))
        style = parse_style(v.get("style", ""))

        # Skip ellipses entirely — at slide scale they're 60px circles crammed together
        # and create an unreadable cluster. Also skip empty-label shapes < threshold.
        if "ellipse" in style:
            skipped_count += 1
            continue
        if w * h < MIN_SHAPE_AREA and not label.strip():
            skipped_count += 1
            continue

        vid = (obj_parent.get("id", "") if obj_parent is not None else "") or v.get("id", "")
        label = resolve_label(v, obj_parent)

        # Detect text-only elements (annotations, notes)
        is_text_only = "text" in style and style.get("fillColor", "none").lower() in ("none", "")

        # Shape type
        if "ellipse" in style:
            shape_type = MSO_SHAPE.OVAL
        elif style.get("rounded") == "1":
            shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        else:
            shape_type = MSO_SHAPE.RECTANGLE

        # Create shape
        sx, sy, sw, sh = to_emu_x(ax), to_emu_y(ay), to_emu_w(w), to_emu_w(h)
        if is_text_only:
            shp = slide.shapes.add_textbox(sx, sy, sw, sh)
        else:
            shp = slide.shapes.add_shape(shape_type, sx, sy, sw, sh)

        # --- Text ---
        tf = shp.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(3)
        tf.margin_right = Pt(3)
        tf.margin_top = Pt(2)
        tf.margin_bottom = Pt(2)
        tf.auto_size = None  # let PowerPoint handle overflow
        tf.text = label

        # Font size — 8pt base, let auto-shrink handle overflow
        font_size = 8
        if "ellipse" in style:
            font_size = 6

        # Font colour from drawio style
        font_color_str = style.get("fontColor", "#000000")
        fr, fg, fb = 0, 0, 0
        if font_color_str and font_color_str.startswith("#"):
            fr, fg, fb = hex_to_rgb(font_color_str)
        elif font_color_str and "rgb" in font_color_str:
            nums = _re.findall(r"\d+", font_color_str)
            if len(nums) >= 3:
                fr, fg, fb = int(nums[0]), int(nums[1]), int(nums[2])

        # Parse drawio font sizes from the HTML label
        # Drawio C4 labels have: <font style="font-size: 16px"><b>Name</b></font>
        # followed by [Type] and <font style="font-size: 11px" color="#cccccc">Description</font>
        raw_html = ""
        if obj_parent is not None and obj_parent.tag == "object":
            raw_html = obj_parent.get("label", "")
        if not raw_html:
            raw_html = v.get("value", "") or ""

        # Extract font-size hints from the HTML for different text segments
        import re as _re3
        font_sizes_in_html = _re3.findall(r'font-size:\s*(\d+)px', raw_html)

        lines = label.split("\n")
        non_empty = [l.strip() for l in lines if l.strip()]
        tf.clear()

        align_val = style.get("align", "center")

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            # Center-align by default (C4 style)
            if align_val == "left":
                p.alignment = PP_ALIGN.LEFT
            else:
                p.alignment = PP_ALIGN.CENTER

            run = p.add_run()
            run.text = line

            # Size the text based on what segment this is
            # First non-empty line = name (large, bold)
            # Lines starting with [ = type (medium)
            # Other lines = description (smaller, lighter)
            line_idx_in_content = non_empty.index(line) if line in non_empty else i

            if line_idx_in_content == 0:
                # Name line — use drawio's font-size if available, else scale
                name_size = float(font_sizes_in_html[0]) if font_sizes_in_html else 16
                run.font.size = Pt(name_size * scale * 0.75)
                run.font.bold = True
                run.font.color.rgb = RGBColor(fr, fg, fb)
            elif line.startswith("["):
                # Type line — regular size
                run.font.size = Pt(font_size)
                run.font.color.rgb = RGBColor(fr, fg, fb)
            else:
                # Description — smaller, lighter
                desc_size = float(font_sizes_in_html[-1]) if len(font_sizes_in_html) > 1 else 11
                run.font.size = Pt(desc_size * scale * 0.75)
                # If base font is light (white), make description dimmer
                if fr > 200 and fg > 200 and fb > 200:
                    run.font.color.rgb = RGBColor(204, 204, 204)
                else:
                    run.font.color.rgb = RGBColor(min(fr + 60, 255), min(fg + 60, 255), min(fb + 60, 255))

        # Vertical text alignment — default to middle unless drawio says top
        va = style.get("verticalAlign", "middle")
        if not is_text_only and hasattr(shp, "text_frame"):
            from pptx.enum.text import MSO_ANCHOR
            if va == "top":
                shp.text_frame.paragraphs[0]  # ensure exists
                # Set anchor via the bodyPr element
                try:
                    from pptx.oxml.ns import qn as _qn2
                    bodyPr = shp.text_frame._txBody.find(_qn2("a:bodyPr"))
                    if bodyPr is not None:
                        bodyPr.set("anchor", "t")
                except Exception:
                    pass
            else:
                try:
                    from pptx.oxml.ns import qn as _qn2
                    bodyPr = shp.text_frame._txBody.find(_qn2("a:bodyPr"))
                    if bodyPr is not None:
                        bodyPr.set("anchor", "ctr")
                except Exception:
                    pass

        if is_text_only:
            shape_map[vid] = shp
            rendered_count += 1
            continue

        # --- Fill ---
        fill_hex = style.get("fillColor", "")
        if fill_hex and fill_hex.lower() != "none":
            if palette:
                fill_hex = snap_colour(fill_hex, palette)
            r, g, b = hex_to_rgb(fill_hex)
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(r, g, b)
        elif fill_hex.lower() == "none":
            shp.fill.background()
        else:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # --- Stroke ---
        stroke_hex = style.get("strokeColor", "")
        if stroke_hex and stroke_hex.startswith("#"):
            if line_darken > 0 and palette:
                stroke_hex = darken(stroke_hex, line_darken)
            r, g, b = hex_to_rgb(stroke_hex)
            shp.line.color.rgb = RGBColor(r, g, b)
        shp.line.width = Pt(0.75)

        if style.get("dashed") == "1":
            shp.line.dash_style = 4

        shape_map[vid] = shp
        rendered_count += 1

    # --- Connectors ---
    # For simple diagrams (few shapes), allow shape-to-shape connectors via edge intersection.
    # For complex diagrams (many shapes), only render edges with explicit mxPoints to avoid noise.
    simple_diagram = rendered_count <= 30
    connector_count = 0
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE

    def _edge_intersect(shp, other_cx, other_cy):
        """Line from shape center to target, clipped at shape rectangle boundary."""
        cx = shp.left + shp.width // 2
        cy = shp.top + shp.height // 2
        ddx = other_cx - cx
        ddy = other_cy - cy
        if ddx == 0 and ddy == 0:
            return cx, cy
        hw = shp.width // 2
        hh = shp.height // 2
        s_x = abs(hw / ddx) if ddx != 0 else float("inf")
        s_y = abs(hh / ddy) if ddy != 0 else float("inf")
        s = min(s_x, s_y)
        return int(cx + ddx * s), int(cy + ddy * s)

    for e, _edge_obj in edges:
        geo = e.find("mxGeometry")
        if geo is None:
            continue

        source_id = e.get("source", "")
        target_id = e.get("target", "")
        src_pt = geo.find("mxPoint[@as='sourcePoint']")
        tgt_pt = geo.find("mxPoint[@as='targetPoint']")

        def _best_edge_pair(s_shp, t_shp):
            """Pick the best source/target edge midpoints for a connector.
            Uses the target's center-x for the source exit point when going vertical,
            so the line goes straight down to the target."""
            sc = (s_shp.left + s_shp.width // 2, s_shp.top + s_shp.height // 2)
            tc = (t_shp.left + t_shp.width // 2, t_shp.top + t_shp.height // 2)
            dx_abs = abs(tc[0] - sc[0])
            dy_abs = abs(tc[1] - sc[1])
            if dy_abs >= dx_abs * 0.5:
                # Primarily vertical — exit from source bottom center, enter target top center
                if sc[1] < tc[1]:
                    # Source exit: use target's x so the line is more vertical
                    src_x = min(max(tc[0], s_shp.left), s_shp.left + s_shp.width)
                    return (src_x, s_shp.top + s_shp.height), (tc[0], t_shp.top)
                else:
                    src_x = min(max(tc[0], s_shp.left), s_shp.left + s_shp.width)
                    return (src_x, s_shp.top), (tc[0], t_shp.top + t_shp.height)
            else:
                # Primarily horizontal
                if sc[0] < tc[0]:
                    return (s_shp.left + s_shp.width, sc[1]), (t_shp.left, tc[1])
                else:
                    return (s_shp.left, sc[1]), (t_shp.left + t_shp.width, tc[1])

        # Resolve endpoints
        if src_pt is not None:
            sx_emu = to_emu_x(float(src_pt.get("x", 0)))
            sy_emu = to_emu_y(float(src_pt.get("y", 0)))
            if tgt_pt is not None:
                tx_emu = to_emu_x(float(tgt_pt.get("x", 0)))
                ty_emu = to_emu_y(float(tgt_pt.get("y", 0)))
            elif simple_diagram and target_id and target_id in shape_map:
                t = shape_map[target_id]
                tx_emu, ty_emu = _edge_intersect(t, sx_emu, sy_emu)
            else:
                continue
        elif simple_diagram and source_id in shape_map:
            s = shape_map[source_id]
            if target_id and target_id in shape_map:
                t = shape_map[target_id]
                (sx_emu, sy_emu), (tx_emu, ty_emu) = _best_edge_pair(s, t)
            elif tgt_pt is not None:
                tx_emu = to_emu_x(float(tgt_pt.get("x", 0)))
                ty_emu = to_emu_y(float(tgt_pt.get("y", 0)))
                sx_emu, sy_emu = _edge_intersect(s, tx_emu, ty_emu)
            else:
                continue
        else:
            continue

        edge_style = parse_style(e.get("style", ""))

        # Use elbow (right-angle) connector if drawio specifies orthogonal routing
        is_orthogonal = "orthogonal" in edge_style.get("edgeStyle", "")
        conn_type = MSO_CONNECTOR_TYPE.ELBOW if is_orthogonal else MSO_CONNECTOR_TYPE.STRAIGHT
        conn = slide.shapes.add_connector(
            conn_type,
            sx_emu, sy_emu, tx_emu, ty_emu,
        )
        stroke = edge_style.get("strokeColor", "")
        if stroke and stroke.startswith("#"):
            er, eg, eb = hex_to_rgb(stroke)
        else:
            er, eg, eb = 120, 120, 120
        conn.line.color.rgb = RGBColor(er, eg, eb)
        stroke_width = edge_style.get("strokeWidth", "1")
        conn.line.width = Pt(float(stroke_width))

        if "dashed" in edge_style or edge_style.get("dashed") == "1":
            conn.line.dash_style = 4

        # Arrowhead — use python-pptx's line end properties
        end_arrow = edge_style.get("endArrow", "")
        end_fill = edge_style.get("endFill", "0")
        if end_arrow and end_arrow != "none":
            try:
                from pptx.oxml.ns import qn as _qn
                from lxml import etree as _lx
                # Navigate to the line element within the connector
                cxnSp = conn._element
                spPr = cxnSp.find(_qn("p:spPr"))
                if spPr is None:
                    spPr = cxnSp.find(_qn("a:spPr"))
                if spPr is not None:
                    ln = spPr.find(_qn("a:ln"))
                    if ln is None:
                        ln = _lx.SubElement(spPr, _qn("a:ln"))
                    # Remove existing tailEnd if any
                    existing = ln.find(_qn("a:tailEnd"))
                    if existing is not None:
                        ln.remove(existing)
                    tail = _lx.SubElement(ln, _qn("a:tailEnd"))
                    tail.set("type", "triangle")
                    tail.set("w", "sm")
                    tail.set("len", "sm")
            except Exception:
                pass

        # Edge label — check the <object> parent for label + placeholder resolution
        edge_obj = _edge_obj  # the object wrapper if present
        edge_val = ""
        if edge_obj is not None and edge_obj.tag == "object":
            edge_val = edge_obj.get("label", "")
        if not edge_val:
            edge_val = e.get("value", "")
        if edge_val:
            lbl = strip_html(edge_val)
            # Resolve %placeholder% tokens from the object parent
            if "%" in lbl and edge_obj is not None:
                lbl = _re.sub(r"%(\w+)%", lambda m: edge_obj.get(m.group(1), m.group(1)), lbl)
            lbl = lbl.strip()
            if lbl:
                mid_x = (sx_emu + tx_emu) // 2
                mid_y = (sy_emu + ty_emu) // 2
                dx_line = abs(tx_emu - sx_emu)
                dy_line = abs(ty_emu - sy_emu)

                if dy_line > dx_line:
                    # Mostly vertical — label to the left of the line
                    # Check if midpoint overlaps any shape, and shift further if so
                    label_w = Pt(max(60, len(lbl) * 4))
                    label_h = Pt(20)
                    lbl_x = mid_x - label_w - Pt(8)
                    lbl_y = mid_y - label_h // 2
                    # Extra shift if label overlaps a shape
                    for _vid, _shp in shape_map.items():
                        if (lbl_x < _shp.left + _shp.width and lbl_x + label_w > _shp.left and
                            lbl_y < _shp.top + _shp.height and lbl_y + label_h > _shp.top):
                            lbl_x = _shp.left - label_w - Pt(8)
                            break
                else:
                    # Mostly horizontal — label centered above the connector
                    # Constrain width to the gap between endpoints
                    gap = dx_line
                    label_w = min(Pt(max(60, len(lbl) * 4)), int(gap * 0.9))
                    label_h = Pt(24)  # taller to allow 2 lines if squeezed
                    lbl_x = mid_x - label_w // 2
                    lbl_y = mid_y - label_h - Pt(2)
                txbox = slide.shapes.add_textbox(lbl_x, lbl_y, label_w, label_h)
                txbox.text_frame.word_wrap = True
                txbox.text_frame.text = lbl
                from pptx.enum.text import PP_ALIGN as _PP
                for p in txbox.text_frame.paragraphs:
                    p.alignment = _PP.CENTER
                    for r in p.runs:
                        r.font.size = Pt(7)
                        r.font.color.rgb = RGBColor(80, 80, 80)

        connector_count += 1

    prs.save(pptx_path)
    print(f"Saved {pptx_path}  ({rendered_count} shapes, {connector_count} connectors, {skipped_count} skipped)")


# ---------------------------------------------------------------------------
# pptx → drawio
# ---------------------------------------------------------------------------

EMU_TO_DRAWIO_PX = 1.0 / DRAWIO_PX_TO_EMU  # inverse


def _emu_to_dx(val: int) -> float:
    return round(val * EMU_TO_DRAWIO_PX, 1)


def _pptx_color_hex(color_format) -> str:
    try:
        rgb = color_format.rgb
        if rgb:
            return f"#{rgb}"
    except Exception:
        pass
    return ""


def _shape_to_style(shp) -> str:
    props: dict[str, str] = {}

    # Shape type
    try:
        auto_type = shp.auto_shape_type
        if auto_type == MSO_SHAPE.OVAL:
            props["ellipse"] = ""
        elif auto_type == MSO_SHAPE.ROUNDED_RECTANGLE:
            props["rounded"] = "1"
    except Exception:
        pass

    props["whiteSpace"] = "wrap"
    props["html"] = "1"

    # Fill
    try:
        if shp.fill.type is not None:
            hex_c = _pptx_color_hex(shp.fill.fore_color)
            if hex_c:
                props["fillColor"] = hex_c
    except Exception:
        props["fillColor"] = "none"

    # Stroke
    try:
        hex_c = _pptx_color_hex(shp.line.color)
        if hex_c:
            props["strokeColor"] = hex_c
    except Exception:
        pass

    return build_style(props)


def pptx_to_drawio(pptx_path: str, drawio_path: str) -> None:
    prs = Presentation(pptx_path)

    mxfile = ET.Element("mxfile")
    mxfile.set("host", "drawio_ppt.py")

    for slide_idx, slide in enumerate(prs.slides):
        diagram = ET.SubElement(mxfile, "diagram")
        diagram.set("name", f"Slide {slide_idx + 1}")
        diagram.set("id", f"slide_{slide_idx}")

        mg = ET.SubElement(diagram, "mxGraphModel")
        mg.set("dx", "0")
        mg.set("dy", "0")
        mg.set("grid", "1")
        mg.set("gridSize", "10")
        mg.set("guides", "1")
        mg.set("tooltips", "1")
        mg.set("connect", "1")
        mg.set("arrows", "1")
        mg.set("fold", "1")
        mg.set("page", "1")
        mg.set("pageScale", "1")
        mg.set("pageWidth", str(int(_emu_to_dx(prs.slide_width))))
        mg.set("pageHeight", str(int(_emu_to_dx(prs.slide_height))))

        root = ET.SubElement(mg, "root")
        ET.SubElement(root, "mxCell", id="0")
        ET.SubElement(root, "mxCell", id="1", parent="0")

        shape_id_map: dict[int, str] = {}  # pptx shape.shape_id → drawio id
        next_id = 2

        for shp in slide.shapes:
            did = str(next_id)
            next_id += 1
            shape_id_map[shp.shape_id] = did

            if shp.shape_type == MSO_SHAPE_TYPE.FREEFORM or shp.shape_type == MSO_SHAPE_TYPE.GROUP:
                continue  # skip complex shapes

            # Check if it's a connector
            if hasattr(shp, "begin_x") and hasattr(shp, "end_x"):
                cell = ET.SubElement(root, "mxCell")
                cell.set("id", did)
                cell.set("value", "")
                cell.set("edge", "1")
                cell.set("parent", "1")
                cell.set("style", "endArrow=classic;html=1;")

                geo = ET.SubElement(cell, "mxGeometry")
                geo.set("relative", "1")
                geo.set("as", "geometry")

                src = ET.SubElement(geo, "mxPoint")
                src.set("x", str(_emu_to_dx(shp.begin_x)))
                src.set("y", str(_emu_to_dx(shp.begin_y)))
                src.set("as", "sourcePoint")

                tgt = ET.SubElement(geo, "mxPoint")
                tgt.set("x", str(_emu_to_dx(shp.end_x)))
                tgt.set("y", str(_emu_to_dx(shp.end_y)))
                tgt.set("as", "targetPoint")
                continue

            # Regular shape (vertex)
            label = ""
            if shp.has_text_frame:
                label = shp.text_frame.text

            cell = ET.SubElement(root, "mxCell")
            cell.set("id", did)
            cell.set("value", label)
            cell.set("style", _shape_to_style(shp))
            cell.set("vertex", "1")
            cell.set("parent", "1")

            geo = ET.SubElement(cell, "mxGeometry")
            geo.set("x", str(_emu_to_dx(shp.left)))
            geo.set("y", str(_emu_to_dx(shp.top)))
            geo.set("width", str(_emu_to_dx(shp.width)))
            geo.set("height", str(_emu_to_dx(shp.height)))
            geo.set("as", "geometry")

    tree = ET.ElementTree(mxfile)
    ET.indent(tree, space="  ")
    tree.write(drawio_path, encoding="unicode", xml_declaration=True)
    total_shapes = sum(len(s.shapes) for s in prs.slides)
    print(f"Saved {drawio_path}  ({len(prs.slides)} slides, {total_shapes} shapes)")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Bidirectional Draw.io ↔ PowerPoint converter",
    )
    sub = parser.add_subparsers(dest="command")

    d2p = sub.add_parser("drawio2ppt", help="Convert .drawio XML to .pptx")
    d2p.add_argument("input", help="Path to .drawio or .drawio.xml file")
    d2p.add_argument("output", help="Output .pptx path")
    d2p.add_argument("--config", help="Path to config.ini with [colours] palette")
    d2p.add_argument("--exclude-layers", nargs="*", default=[], help="Layer names to exclude")
    d2p.add_argument("--back-layers", nargs="*", default=[], help="Layer names to send to back")
    d2p.add_argument("--line-darken", type=int, default=20, help="Darken stroke by N (default 20)")

    p2d = sub.add_parser("ppt2drawio", help="Convert .pptx to .drawio")
    p2d.add_argument("input", help="Path to .pptx file")
    p2d.add_argument("output", help="Output .drawio path")

    args = parser.parse_args()

    if args.command == "drawio2ppt":
        drawio_to_pptx(
            args.input, args.output,
            config_path=args.config,
            exclude_layers=args.exclude_layers,
            back_layers=args.back_layers,
            line_darken=args.line_darken,
        )
    elif args.command == "ppt2drawio":
        pptx_to_drawio(args.input, args.output)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
