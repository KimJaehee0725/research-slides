#!/usr/bin/env python3
"""Get figures and tables out of a paper PDF as clean PNGs.

    # 1) find captions and their page/bbox
    python3 scripts/crop_figure.py paper.pdf --list
    # 2) look at a page (writes page PNG with a 50pt grid so bboxes are easy to read)
    python3 scripts/crop_figure.py paper.pdf --page 4 --grid --out p4.png
    # 3a) auto-crop a captioned figure (heuristic: graphics above/below the caption)
    python3 scripts/crop_figure.py paper.pdf --auto "Figure 2" --out figs/fig2.png
    # 3b) or crop an exact region in PDF points (x0 y0 x1 y1, origin top-left)
    python3 scripts/crop_figure.py paper.pdf --page 4 --bbox 60 80 550 330 --out figs/fig2.png

Always open the PNG and check it before using it: auto-crop can miss
axis labels or grab body text. Default 250 dpi keeps figures sharp on a
20-inch slide. The caption itself is excluded; put it in the band text.
"""
import argparse
import re
import sys
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    try:
        import fitz
    except ImportError:
        sys.exit("PyMuPDF is required: pip install pymupdf")

CAP = re.compile(r"^\s*(Figure|Fig\.|Table)\s*(\d+)\s*[\.:]", re.I)  # real captions only


def captions(doc):
    """First caption block per label: (label, page, caption bbox, text)."""
    out, seen = [], set()
    for pno, page in enumerate(doc):
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[:5]
            m = CAP.match(txt)
            if m:
                kind = "Table" if m.group(1).lower() == "table" else "Figure"
                if f"{kind} {m.group(2)}" in seen:
                    continue
                seen.add(f"{kind} {m.group(2)}")
                out.append((f"{kind} {m.group(2)}", pno + 1, (x0, y0, x1, y1),
                            " ".join(txt.split())[:90]))
    return out


def graphics_rects(page):
    rects = [fitz.Rect(d["rect"]) for d in page.get_drawings() if d["rect"].width > 2 or d["rect"].height > 2]
    for img in page.get_images(full=True):
        for r in page.get_image_rects(img[0]):
            rects.append(fitz.Rect(r))
    return rects


def auto_bbox(doc, label):
    caps = [c for c in captions(doc) if c[0].lower() == label.lower()]
    if not caps:
        sys.exit(f"caption '{label}' not found; run --list")
    _, pno, cb, _ = caps[0]
    page = doc[pno - 1]
    cap = fitz.Rect(cb)
    is_table = label.lower().startswith("table")
    text_blocks = [fitz.Rect(b[:4]) for b in page.get_text("blocks")
                   if not CAP.match(b[4]) and len(b[4].split()) > 25]
    g = graphics_rects(page)
    # figures: graphics above the caption; tables: below (captions usually on top)
    if is_table:
        cand = [r for r in g if r.y0 >= cap.y1 - 2]
        limit = min([t.y0 for t in text_blocks if t.y0 > cap.y1] + [page.rect.y1])
        cand = [r for r in cand if r.y1 <= limit + 2]
        words = [fitz.Rect(w[:4]) for w in page.get_text("words") if cap.y1 < w[1] < limit]
        cand += words
    else:
        limit = max([t.y1 for t in text_blocks if t.y1 < cap.y0] + [page.rect.y0])
        cand = [r for r in g if r.y1 <= cap.y0 + 2 and r.y0 >= limit - 2]
        words = [fitz.Rect(w[:4]) for w in page.get_text("words") if limit < w[1] and w[3] < cap.y0]
        cand += words
    cand = [r for r in cand if r.x0 < cap.x1 + 40 and r.x1 > cap.x0 - 40] or cand
    if not cand:
        sys.exit("no graphics found near caption; use --page/--bbox")
    bb = cand[0]
    for r in cand[1:]:
        bb |= r
    return pno, bb + (-4, -4, 4, 4)


def render(doc, pno, bbox, out, dpi, grid=False):
    page = doc[pno - 1]
    clip = fitz.Rect(bbox) if bbox else page.rect
    pix = page.get_pixmap(dpi=dpi, clip=clip, alpha=False)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    pix.save(str(out))
    if grid:
        from PIL import Image, ImageDraw
        im = Image.open(out)
        d = ImageDraw.Draw(im)
        s = dpi / 72
        for v in range(0, int(clip.width) + 1, 50):
            d.line([(v * s, 0), (v * s, im.height)], fill=(255, 0, 0), width=1)
            d.text((v * s + 2, 2), str(int(clip.x0 + v)), fill=(255, 0, 0))
        for v in range(0, int(clip.height) + 1, 50):
            d.line([(0, v * s), (im.width, v * s)], fill=(0, 0, 255), width=1)
            d.text((2, v * s + 2), str(int(clip.y0 + v)), fill=(0, 0, 255))
        im.save(out)
    print(f"page {pno} bbox {tuple(round(v, 1) for v in clip)} -> {out} ({pix.width}x{pix.height})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("--list", action="store_true",
                    help="list captions: label, page, CAPTION bbox (not the figure bbox), text")
    ap.add_argument("--page", type=int, help="1-based page to render (whole page, or --bbox region)")
    ap.add_argument("--bbox", type=float, nargs=4, metavar=("X0", "Y0", "X1", "Y1"),
                    help="region in PDF points, origin top-left (read it off a --grid render)")
    ap.add_argument("--auto", metavar="LABEL",
                    help='auto-crop the figure/table whose caption starts with LABEL, e.g. "Figure 2"')
    ap.add_argument("--out", default="figure.png", help="output PNG path")
    ap.add_argument("--dpi", type=int, default=250, help="output resolution (grid renders cap at 110)")
    ap.add_argument("--grid", action="store_true",
                    help="with --page: overlay a 50-pt grid labelled in PDF points")
    a = ap.parse_args()
    doc = fitz.open(a.pdf)
    if a.list:
        for label, pno, bb, txt in captions(doc):
            print(f"{label:10s} p{pno:<3d} caption_bbox={tuple(round(v) for v in bb)}  {txt}")
    elif a.auto:
        pno, bb = auto_bbox(doc, a.auto)
        render(doc, pno, bb, a.out, a.dpi)
    elif a.page:
        render(doc, a.page, a.bbox, a.out, a.dpi if not a.grid else min(a.dpi, 110), grid=a.grid)
    else:
        ap.print_help()
