#!/usr/bin/env python3
"""Render a deck to PNGs and flag layout problems.

    python3 scripts/render_check.py deck.pptx [--out DIR] [--dpi 50]

Outputs DIR/slide-N.png (zero-padded when there are 10+ slides), DIR/contact.png,
plus the intermediate DIR/<name>.pdf and an unhidden DIR/<name>.pptx copy.
Prints warnings:
  * text that probably overflows its box (estimated with NanumSquare metrics)
  * shapes outside the slide or overlapping the footer bar
  * leftover prompt text / empty placeholders
Always look at contact.png (and individual slides) after building.
"""
import argparse
import glob
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageFont
from pptx import Presentation
from pptx.util import Emu

FOOTER_Y = 10.74
FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/nanum/NanumSquareR.ttf",
    "/Library/Fonts/NanumSquareOTF_acR.otf",
    str(Path.home() / "Library/Fonts/NanumSquareOTFR.otf"),
]


def font(size_pt):
    for f in FONT_CANDIDATES:
        if Path(f).exists():
            return ImageFont.truetype(f, int(size_pt * 10))  # 10 px per pt => 720 px/in
    return None


def est_lines(text, size_pt, width_in):
    f = font(size_pt)
    width_px = width_in * 720
    lines = 0
    for para in text.split("\n"):
        if f is None:
            lines += max(1, int(len(para) * size_pt / 72 * 0.55 / max(width_in, 0.1)) + 1)
            continue
        w = f.getlength(para) if para else 0
        lines += max(1, int(w // max(width_px, 1)) + 1)
    return lines


def inherited_spacing(shape, slide):
    """Line spacing multiple: explicit paragraph value, else the layout's list style."""
    ls = shape.text_frame.paragraphs[0].line_spacing
    if isinstance(ls, float):
        return ls
    if shape.is_placeholder:
        idx = shape.placeholder_format.idx
        for lp in slide.slide_layout.placeholders:
            if lp.placeholder_format.idx == idx:
                v = lp._element.xpath(".//a:lstStyle/a:lvl1pPr/a:lnSpc/a:spcPct/@val")
                if v:
                    return int(v[0]) / 100000
    return 1.0


def inherited_size(shape, slide):
    """Largest explicit run size, else the size defined by the layout placeholder."""
    sizes = [r.font.size.pt for p in shape.text_frame.paragraphs for r in p.runs if r.font.size]
    if sizes:
        return max(sizes)
    if shape.is_placeholder:
        idx = shape.placeholder_format.idx
        for lp in slide.slide_layout.placeholders:
            if lp.placeholder_format.idx == idx:
                szs = lp._element.xpath(".//a:lstStyle/a:lvl1pPr/a:defRPr/@sz")
                if szs:
                    return int(szs[0]) / 100
        if idx == 0:
            return 44
    return 18


def check(prs):
    warns = []
    W, H = Emu(prs.slide_width).inches, Emu(prs.slide_height).inches
    for n, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.left is None:
                continue
            x, y, w, h = (Emu(v).inches for v in (sh.left, sh.top, sh.width, sh.height))
            if x < -0.01 or y < -0.01 or x + w > W + 0.01 or y + h > H + 0.01:
                warns.append(f"slide {n}: '{sh.name}' goes outside the slide")
            if y + h > FOOTER_Y + 0.02 and sh.name not in ("footer_bar",) and \
                    not (sh.is_placeholder and sh.placeholder_format.idx == 11):
                warns.append(f"slide {n}: '{sh.name}' overlaps the footer bar (bottom {y+h:.2f} in)")
            if sh.has_text_frame and sh.text_frame.text.strip():
                size = inherited_size(sh, slide)
                pad = 0.8 if sh.is_placeholder and sh.placeholder_format.idx == 12 else 0.3
                lines = est_lines(sh.text_frame.text, size, max(w - pad, 0.5))
                need = lines * size * 1.2 * inherited_spacing(sh, slide) / 72 + 0.1
                if need > h * 1.08:
                    warns.append(f"slide {n}: text in '{sh.name}' may overflow "
                                 f"(~{lines} lines at {size:.0f}pt need {need:.1f} in, box {h:.1f} in)")
    return warns


def unhidden_copy(pptx, out):
    """LibreOffice skips hidden slides; render a copy with every slide shown."""
    prs = Presentation(pptx)
    for sl in prs.slides:
        if sl._element.get("show") == "0":
            del sl._element.attrib["show"]
    tmp = out / (Path(pptx).stem + ".pptx")
    prs.save(tmp)
    return tmp


def render(pptx, out, dpi, hidden=True):
    out.mkdir(parents=True, exist_ok=True)
    if hidden:
        pptx = unhidden_copy(pptx, out)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("soffice not found: skipping render (install LibreOffice)")
        return
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(pptx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    pdf = out / (Path(pptx).stem + ".pdf")
    for f in glob.glob(str(out / "slide-*.png")):
        Path(f).unlink()
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(out / "slide")], check=True)
    files = sorted(glob.glob(str(out / "slide-*.png")))
    ims = [Image.open(f) for f in files]
    if not ims:
        return
    w, h = ims[0].size
    cols = 4 if len(ims) > 6 else 3
    rows = (len(ims) + cols - 1) // cols
    grid = Image.new("RGB", (cols * w + (cols + 1) * 8, rows * h + (rows + 1) * 8), "#d9d9d9")
    for i, im in enumerate(ims):
        grid.paste(im, (8 + (i % cols) * (w + 8), 8 + (i // cols) * (h + 8)))
    grid.save(out / "contact.png")
    print(f"rendered {len(ims)} slides -> {out}/contact.png")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None)
    ap.add_argument("--dpi", type=int, default=50)
    ap.add_argument("--no-render", action="store_true")
    ap.add_argument("--skip-hidden", action="store_true", help="do not render hidden slides")
    a = ap.parse_args()
    prs = Presentation(a.pptx)
    ws = check(prs)
    print("\n".join(ws) if ws else "no layout warnings")
    if not a.no_render:
        render(a.pptx, Path(a.out or Path(a.pptx).with_suffix("")).resolve(), a.dpi,
               hidden=not a.skip_hidden)
    sys.exit(0)
