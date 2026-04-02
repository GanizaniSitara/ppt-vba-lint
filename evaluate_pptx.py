"""
evaluate_pptx.py — Capture a PowerPoint slide as PNG and compare against a reference.

Produces:
  - Side-by-side comparison image
  - Amplified diff image
  - Scorecard with pixel metrics

Usage:
    python evaluate_pptx.py <pptx_path> <reference_png> [--iteration N]
"""

import sys
import os
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFont, ImageOps
from pptx import Presentation
from pptx.util import Emu


def pptx_slide_to_image(pptx_path: str, slide_idx: int = 0, scale: float = 2.0) -> Image.Image:
    """Render a pptx slide to a PIL Image by drawing shapes with PIL.
    This is a lightweight renderer — not pixel-perfect but good enough for delta comparison."""
    import re

    prs = Presentation(pptx_path)
    slide = prs.slides[slide_idx]
    sw = prs.slide_width
    sh = prs.slide_height

    # Target image size
    img_w = int(sw / 914400 * 96 * scale)
    img_h = int(sh / 914400 * 96 * scale)
    img = Image.new("RGB", (img_w, img_h), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    def emu_to_px(emu):
        return int(emu / 914400 * 96 * scale)

    try:
        font = ImageFont.truetype("arial.ttf", int(6 * scale))
        font_sm = ImageFont.truetype("arial.ttf", int(4 * scale))
    except Exception:
        font = ImageFont.load_default()
        font_sm = font

    for shp in slide.shapes:
        x = emu_to_px(shp.left)
        y = emu_to_px(shp.top)
        w = emu_to_px(shp.width)
        h = emu_to_px(shp.height)

        # Get colours
        fill_color = (255, 255, 255)
        outline_color = (0, 0, 0)

        try:
            if shp.fill.type is not None:
                rgb = shp.fill.fore_color.rgb
                fill_color = (rgb[0], rgb[1], rgb[2])
        except Exception:
            pass

        try:
            rgb = shp.line.color.rgb
            outline_color = (rgb[0], rgb[1], rgb[2])
        except Exception:
            pass

        # Determine shape type
        is_oval = False
        is_rounded = False
        is_connector = False

        try:
            from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
            if shp.shape_type == MSO_SHAPE_TYPE.FREEFORM:
                is_connector = True
            auto = shp.auto_shape_type
            if auto == MSO_SHAPE.OVAL:
                is_oval = True
            elif auto == MSO_SHAPE.ROUNDED_RECTANGLE:
                is_rounded = True
        except Exception:
            pass

        # Check if it's a connector (has begin_x/end_x)
        if hasattr(shp, "begin_x") and hasattr(shp, "end_x"):
            bx = emu_to_px(shp.begin_x)
            by = emu_to_px(shp.begin_y)
            ex = emu_to_px(shp.end_x)
            ey = emu_to_px(shp.end_y)
            draw.line([(bx, by), (ex, ey)], fill=outline_color, width=max(1, int(scale * 0.5)))
            continue

        if is_connector:
            continue

        # Draw shape
        if is_oval:
            draw.ellipse([x, y, x + w, y + h], fill=fill_color, outline=outline_color, width=1)
        elif is_rounded:
            r = max(2, int(min(w, h) * 0.08))
            draw.rounded_rectangle([x, y, x + w, y + h], radius=r, fill=fill_color, outline=outline_color, width=1)
        else:
            draw.rectangle([x, y, x + w, y + h], fill=fill_color, outline=outline_color, width=1)

        # Text
        if shp.has_text_frame:
            text = shp.text_frame.text.strip()
            if text:
                f = font_sm if is_oval else font
                # Truncate to fit
                max_chars = max(5, w // int(4 * scale))
                lines = text.split("\n")
                ty = y + int(2 * scale)
                for line in lines[:5]:
                    draw.text((x + int(3 * scale), ty), line[:max_chars], fill=(0, 0, 0), font=f)
                    ty += int(8 * scale)

    return img


def add_label(img, text, bg_color=(30, 30, 30)):
    label_h = 40
    out = Image.new("RGB", (img.width, img.height + label_h), bg_color)
    d = ImageDraw.Draw(out)
    try:
        font = ImageFont.truetype("arial.ttf", 24)
    except Exception:
        font = ImageFont.load_default()
    bbox = d.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    d.text(((img.width - tw) // 2, 8), text, fill="white", font=font)
    out.paste(img, (0, label_h))
    return out


def compare(img_a: Image.Image, img_b: Image.Image, tag_a: str, tag_b: str, out_dir: str, iteration: int):
    """Generate comparison images and pixel metrics."""
    # Resize to same dimensions
    w = max(img_a.width, img_b.width)
    h = max(img_a.height, img_b.height)
    a = Image.new("RGB", (w, h), (255, 255, 255))
    b = Image.new("RGB", (w, h), (255, 255, 255))
    a.paste(img_a, (0, 0))
    b.paste(img_b, (0, 0))

    # Side-by-side
    la = add_label(a, tag_a)
    lb = add_label(b, tag_b)
    gap = 6
    sbs = Image.new("RGB", (la.width + lb.width + gap, la.height), (40, 40, 40))
    sbs.paste(la, (0, 0))
    sbs.paste(lb, (la.width + gap, 0))
    sbs_path = os.path.join(out_dir, f"iter{iteration:02d}_SBS.png")
    sbs.save(sbs_path)

    # Amplified diff
    diff = ImageChops.difference(a, b)
    diff = ImageEnhance.Brightness(diff).enhance(4.0)
    diff = add_label(diff, f"DIFF iter{iteration}: {tag_a} vs {tag_b}")
    diff_path = os.path.join(out_dir, f"iter{iteration:02d}_DIFF.png")
    diff.save(diff_path)

    # Pixel metrics
    import numpy as np
    arr_a = np.array(a).astype(float)
    arr_b = np.array(b).astype(float)
    diff_arr = np.abs(arr_a - arr_b)
    pct_identical = (diff_arr < 5).all(axis=2).mean() * 100
    pct_close = (diff_arr < 20).all(axis=2).mean() * 100
    mad = diff_arr.mean()

    return {
        "iteration": iteration,
        "pct_identical": pct_identical,
        "pct_close": pct_close,
        "mad": mad,
        "sbs_path": sbs_path,
        "diff_path": diff_path,
    }


def print_scorecard(results: list[dict]):
    print()
    print("=" * 60)
    print("QUALITY SCORECARD")
    print("=" * 60)
    print(f"{'Iter':>5} | {'Identical':>10} | {'Close':>8} | {'MAD':>8}")
    print("-" * 60)
    for r in results:
        print(f"  {r['iteration']:>3} | {r['pct_identical']:>8.1f}% | {r['pct_close']:>6.1f}% | {r['mad']:>6.1f}")
    print("=" * 60)
    if len(results) > 1:
        d = results[-1]["pct_close"] - results[0]["pct_close"]
        print(f"DELTA from iter {results[0]['iteration']} → {results[-1]['iteration']}: {d:+.1f}% close pixels")


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx_path")
    parser.add_argument("reference_png")
    parser.add_argument("--iteration", type=int, default=0)
    parser.add_argument("--out-dir", default="C:/temp/pptx_eval")
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    print(f"Rendering slide from {args.pptx_path}...")
    pptx_img = pptx_slide_to_image(args.pptx_path, scale=2.0)
    pptx_img.save(os.path.join(args.out_dir, f"iter{args.iteration:02d}_pptx.png"))

    print(f"Loading reference {args.reference_png}...")
    ref_img = Image.open(args.reference_png).convert("RGB")
    # Resize reference to match pptx render size
    ref_img = ref_img.resize((pptx_img.width, pptx_img.height), Image.LANCZOS)

    result = compare(pptx_img, ref_img, f"PPTX iter{args.iteration}", "REFERENCE (drawio)", args.out_dir, args.iteration)
    print_scorecard([result])

    return result


if __name__ == "__main__":
    main()
