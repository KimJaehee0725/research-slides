#!/usr/bin/env python3
"""Split a multi-panel figure into its panels (XY-cut on background gaps).

    python3 scripts/split_panels.py figure.png --out figs/fig3_panels [--min-gap 12] [--depth 3]

Writes <out>/p1.png, p2.png, ... , <out>/panels.json (crop fractions l,t,r,b of the
source) and <out>/preview.png (numbered boxes). Always open preview.png: if a panel
was split too finely or not at all, adjust --min-gap / --depth, or write the crops
by hand into the spec (`panels` items accept "crop": [l, t, r, b]).
Pieces that belong together (a card and its caption, a heatmap and its colour bar)
can be joined with --merge "2+5,3+6"; titles removed with --drop 1.

Why: a figure laid out for a paper page (e.g. 3 + 2 panels, or tall stacks) rarely fits
a 16:9 content zone. Cut it into panels and let build_deck's `panels` shape re-flow them
into rows that fill the zone (references/spec.md, `panels`).
"""
import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw


def bg_mask(a, tol):
    """True where a pixel looks like background (close to the corner colour or white)."""
    corners = np.stack([a[0, 0], a[0, -1], a[-1, 0], a[-1, -1]]).astype(int)
    bg = np.median(corners, axis=0)
    d = np.abs(a.astype(int) - bg).max(axis=2)
    white = (a.astype(int) > 245).all(axis=2)
    return (d <= tol) | white


def runs(flags, min_len):
    """Index ranges [s, e) where flags is True for at least min_len in a row."""
    out, s = [], None
    for i, f in enumerate(list(flags) + [False]):
        if f and s is None:
            s = i
        elif not f and s is not None:
            if i - s >= min_len:
                out.append((s, i))
            s = None
    return out


def cut(mask, box, axis, min_gap, depth, min_size):
    x0, y0, x1, y1 = box
    sub = mask[y0:y1, x0:x1]
    if sub.size == 0:
        return []
    # trim background margins
    rows_bg, cols_bg = sub.all(axis=1), sub.all(axis=0)
    if rows_bg.all():
        return []
    ty = np.argmax(~rows_bg); by = len(rows_bg) - np.argmax(~rows_bg[::-1])
    tx = np.argmax(~cols_bg); bx = len(cols_bg) - np.argmax(~cols_bg[::-1])
    x0, y0, x1, y1 = x0 + tx, y0 + ty, x0 + bx, y0 + by
    if depth == 0:
        return [(x0, y0, x1, y1)]
    sub = mask[y0:y1, x0:x1]
    for ax in (axis, 1 - axis):
        flags = sub.all(axis=1) if ax == 0 else sub.all(axis=0)   # ax 0: horizontal gaps
        gaps = runs(flags, min_gap)
        if gaps:
            edges = [0] + [g for gap in gaps for g in gap] + [len(flags)]
            parts = []
            for s, e in zip(edges[::2], edges[1::2]):
                if e - s < min_size:
                    continue
                b = (x0, y0 + s, x1, y0 + e) if ax == 0 else (x0 + s, y0, x0 + e, y1)
                parts += cut(mask, b, 1 - ax, min_gap, depth - 1, min_size)
            if len(parts) > 1:
                return parts
    return [(x0, y0, x1, y1)]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("image")
    ap.add_argument("--out", required=True, help="output directory")
    ap.add_argument("--min-gap", type=int, default=12, help="min background gap (px) between panels")
    ap.add_argument("--depth", type=int, default=3, help="max nesting of cuts")
    ap.add_argument("--tol", type=int, default=6, help="colour tolerance for background")
    ap.add_argument("--pad", type=int, default=6, help="padding (px) kept around each panel")
    ap.add_argument("--merge", default="", help='join over-split pieces, e.g. "2+5,3+6" (numbers from a first run)')
    ap.add_argument("--drop", default="", help='discard pieces, e.g. "1" (a figure title)')
    a = ap.parse_args()

    im = Image.open(a.image).convert("RGB")
    arr = np.asarray(im)
    mask = bg_mask(arr, a.tol)
    W, H = im.size
    boxes = cut(mask, (0, 0, W, H), 0, a.min_gap, a.depth, max(20, a.min_gap * 2))
    boxes.sort(key=lambda b: (round(b[1] / (H * 0.05)), b[0]))  # reading order
    if a.merge or a.drop:  # numbers refer to the reading order of the plain split
        groups = [[int(v) - 1 for v in g.split("+")] for g in a.merge.split(",") if g]
        drop = {int(v) - 1 for v in a.drop.split(",") if v}
        used = {i for g in groups for i in g} | drop
        merged = [(min(boxes[i][0] for i in g), min(boxes[i][1] for i in g),
                   max(boxes[i][2] for i in g), max(boxes[i][3] for i in g)) for g in groups]
        boxes = merged + [b for i, b in enumerate(boxes) if i not in used]
        boxes.sort(key=lambda b: (round(b[1] / (H * 0.05)), b[0]))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    info = []
    prev = im.copy()
    d = ImageDraw.Draw(prev)
    for i, (x0, y0, x1, y1) in enumerate([tuple(int(v) for v in b) for b in boxes], 1):
        x0, y0 = max(0, x0 - a.pad), max(0, y0 - a.pad)
        x1, y1 = min(W, x1 + a.pad), min(H, y1 + a.pad)
        im.crop((x0, y0, x1, y1)).save(out / f"p{i}.png")
        crop = [round(x0 / W, 4), round(y0 / H, 4), round(1 - x1 / W, 4), round(1 - y1 / H, 4)]
        info.append({"file": f"p{i}.png", "crop": crop, "px": [x0, y0, x1, y1],
                     "aspect": round((x1 - x0) / (y1 - y0), 2)})
        d.rectangle((x0, y0, x1, y1), outline=(220, 30, 30), width=4)
        d.rectangle((x0, y0, x0 + 46, y0 + 34), fill=(220, 30, 30))
        d.text((x0 + 10, y0 + 6), str(i), fill="white")
    prev.save(out / "preview.png")
    (out / "panels.json").write_text(json.dumps({"source": str(a.image), "panels": info}, indent=2))
    for p in info:
        print(p["file"], "crop", p["crop"], "aspect", p["aspect"])
    print(f"-> {out}/preview.png")


if __name__ == "__main__":
    main()
