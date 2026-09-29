#!/usr/bin/env python3
"""Print every layout and placeholder of a template as Markdown.

    python3 scripts/describe_template.py [template.pptx] > references/layouts_generated.md

Use it on an unfamiliar copy of the template to see which roles exist.
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_template import LAYOUTS  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent


def main(path):
    prs = Presentation(path)
    purpose = {n: p for n, _, _, p in LAYOUTS}
    print(f"슬라이드 크기: {Emu(prs.slide_width).inches:.2f} × {Emu(prs.slide_height).inches:.2f} in\n")
    for lay in prs.slide_layouts:
        print(f"### `{lay.name}`\n")
        if lay.name in purpose:
            print(f"{purpose[lay.name]}\n")
        print("| 역할 | idx | 종류 | x, y, w, h (in) | 글자 크기 |")
        print("|---|---|---|---|---|")
        for p in lay.placeholders:
            pf = p.placeholder_format
            kind = str(pf.type).split(".")[-1].split(" ")[0]
            sz = p._element.xpath(".//a:lstStyle/a:lvl1pPr/a:defRPr/@sz")
            size = f"{int(sz[0])//100}pt" if sz else ("44pt" if pf.idx == 0 else "-")
            box = ", ".join(f"{Emu(v).inches:.2f}" for v in (p.left, p.top, p.width, p.height))
            print(f"| `{p.name}` | {pf.idx} | {kind} | {box} | {size} |")
        print()


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ROOT / "assets" / "research-slides-template.pptx")
