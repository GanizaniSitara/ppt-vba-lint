"""
iterate_c4.py — Run N improvement iterations on the C4 diagram conversion,
capturing red/green overlay + metrics after each.

Usage: python iterate_c4.py
"""
import os
import sys
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageChops, ImageEnhance, ImageOps

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from evaluate_pptx import pptx_slide_to_image

DRAWIO = "C:/Documents/2025-Q1/C4.drawio"
PPTX = "C:/temp/C4_output.pptx"
REF = "C:/temp/C4_reference.png"
OUT = "C:/temp/pptx_eval"
PYTHON = "C:/miniconda3/envs/python312/python.exe"
CONVERTER = "C:/git/ppt-vba-lint/drawio_ppt.py"

os.makedirs(OUT, exist_ok=True)


def add_label(img, text):
    h = 32
    out = Image.new("RGB", (img.width, img.height + h), (30, 30, 30))
    d = ImageDraw.Draw(out)
    try:
        f = ImageFont.truetype("arial.ttf", 20)
    except Exception:
        f = ImageFont.load_default()
    bb = d.textbbox((0, 0), text, font=f)
    d.text(((img.width - (bb[2] - bb[0])) // 2, 6), text, fill="white", font=f)
    out.paste(img, (0, h))
    return out


def evaluate(iteration: int) -> dict:
    """Render PPTX, compare to reference, save overlay + diff, return metrics."""
    pptx_img = pptx_slide_to_image(PPTX, scale=2.0)
    ref = Image.open(REF).convert("RGB")
    ref = ref.resize((pptx_img.width, pptx_img.height), Image.LANCZOS)

    a = Image.new("RGB", (pptx_img.width, pptx_img.height), (255, 255, 255))
    b = Image.new("RGB", (pptx_img.width, pptx_img.height), (255, 255, 255))
    a.paste(pptx_img, (0, 0))
    b.paste(ref, (0, 0))

    # Red/green overlay
    r_ch, g_ch, _ = a.split()
    _, g2, _ = b.split()
    overlay = Image.merge("RGB", [ImageOps.invert(r_ch), ImageOps.invert(g2), Image.new("L", r_ch.size, 0)])
    overlay = add_label(overlay, f"iter {iteration}: PPTX (red) vs REF (green)")
    overlay.save(os.path.join(OUT, f"C4_iter{iteration:02d}_OVERLAY.png"))

    # Diff
    diff = ImageChops.difference(a, b)
    diff = ImageEnhance.Brightness(diff).enhance(4.0)
    diff = add_label(diff, f"iter {iteration}: DIFF")
    diff.save(os.path.join(OUT, f"C4_iter{iteration:02d}_DIFF.png"))

    # Save rendered pptx
    pptx_img.save(os.path.join(OUT, f"C4_iter{iteration:02d}_pptx.png"))

    # Metrics
    arr_a = np.array(a).astype(float)
    arr_b = np.array(b).astype(float)
    d = np.abs(arr_a - arr_b)
    return {
        "iter": iteration,
        "identical": (d < 5).all(axis=2).mean() * 100,
        "close": (d < 20).all(axis=2).mean() * 100,
        "mad": d.mean(),
    }


def convert():
    """Run the converter."""
    import subprocess
    r = subprocess.run(
        [PYTHON, CONVERTER, "drawio2ppt", DRAWIO, PPTX],
        capture_output=True, text=True,
    )
    print(f"  {r.stdout.strip()}")
    if r.returncode != 0:
        print(f"  ERROR: {r.stderr.strip()}")
    return r.returncode == 0


def main():
    results = []

    # Baseline
    print("=== Iteration 0 (baseline) ===")
    convert()
    r = evaluate(0)
    results.append(r)
    print(f"  Identical: {r['identical']:.1f}%  Close: {r['close']:.1f}%  MAD: {r['mad']:.1f}")

    # Print final scorecard
    print()
    print("=" * 65)
    print(f"{'Iter':>5} | {'Identical':>10} | {'Close':>8} | {'MAD':>8}")
    print("-" * 65)
    for r in results:
        print(f"  {r['iter']:>3} | {r['identical']:>8.1f}% | {r['close']:>6.1f}% | {r['mad']:>6.1f}")
    print("=" * 65)


if __name__ == "__main__":
    main()
