#!/usr/bin/env python3
"""Build a deck from a JSON spec using the research-slides template.

    python3 scripts/build_deck.py spec.json out.pptx [--template T.pptx] [--keep-slides]

See references/spec.md for the full schema. Minimal slide:

    {"layout": "A_FigureFull_Band", "section": "Method", "title": "...",
     "figure": "figs/fig2.png", "band": ["주장", "> 근거"], "notes": "내레이션"}

Rules enforced here so decks stay consistent:
  * text goes into named placeholders (roles), never free text boxes
  * figures are fitted (no distortion, no content cropped) inside their role box
  * unused placeholders are removed so no prompt text leaks into the deck
  * every drawn shape goes through components.py (fixed colours per kind)
"""
import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Inches

sys.path.insert(0, str(Path(__file__).resolve().parent))
import components as C  # noqa: E402
from rs_style import EMU  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = ROOT / "assets" / "research-slides-template.pptx"

# Free zones per layout (inches): where "shapes" with frame="zone" are placed.
ZONES = {
    "A_FigureFull_Band": (0.75, 2.25, 18.5, 5.10),
    "B_FigureTwo_Band": (0.75, 2.25, 18.5, 5.10),
    "C_FigureExplain_Band": (0.75, 2.25, 18.5, 5.10),
    "D_TopBand_Wide": (0.75, 4.45, 18.5, 6.00),
    "E_TwoColumn": (0.75, 2.30, 18.5, 7.95),
    "F_Diagram_Band": (0.75, 2.25, 18.5, 5.10),
    "G_TitleOnly": (0.75, 2.25, 18.5, 8.20),
    "H_Statement": (0.75, 2.25, 18.5, 8.20),
}
WARNINGS = []
TEXT_ROLES_WITH_LEVELS = {"band", "explain", "body_left", "body_right", "body"}
CONTENT_LAYOUTS = set(ZONES)


def layout_roles(layout):
    """role name (layout placeholder name) -> placeholder idx."""
    return {p.name: p.placeholder_format.idx for p in layout.placeholders}


def find_layout(prs, name):
    for l in prs.slide_layouts:
        if l.name == name:
            return l
    names = ", ".join(l.name for l in prs.slide_layouts)
    raise SystemExit(f"unknown layout '{name}'. available: {names}")


def clear_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)


def set_footer(prs, text):
    for shp in prs.slide_masters[0].shapes:
        if shp.name == "footer_text":
            C.fill_text(shp.text_frame, text, size=20, color="bg1", levels=False)


def resolve(base, p):
    p = Path(p)
    return p if p.is_absolute() else (base / p)


def zone_box(layout_name, b, frame):
    """Convert a box to inches. frame='zone' -> fractions of the layout free zone."""
    if frame != "zone":
        return b
    zx, zy, zw, zh = ZONES.get(layout_name, (0.75, 2.25, 18.5, 8.2))
    x, y, w, h = b
    return (zx + x * zw, zy + y * zh, w * zw, h * zh)


def place_figure(slide, ph, spec, base):
    """Fit an image into a picture placeholder's box, keeping aspect ratio."""
    if isinstance(spec, str):
        spec = {"path": spec}
    x, y, w, h = (Emu(v).inches for v in (ph.left, ph.top, ph.width, ph.height))
    ph._element.getparent().remove(ph._element)
    pic, (dx, dy, dw, dh) = C.image_fit(slide, resolve(base, spec["path"]), x, y, w, h,
                                        crop=spec.get("crop"), align=spec.get("align", "c"))
    fill_ratio = (dw * dh) / (w * h)
    if fill_ratio < 0.45:
        WARNINGS.append(f"figure '{Path(spec['path']).name}' fills only {fill_ratio:.0%} of its box "
                        f"({dw:.1f}x{dh:.1f} of {w:.1f}x{h:.1f} in): crop to the needed panel or "
                        f"pick a layout whose box matches its aspect ratio (references/layouts.md)")
    for hl in spec.get("highlights", []):
        hx, hy, hw, hh = hl["box"] if isinstance(hl, dict) else hl
        bx, by, bw, bh = dx + hx * dw, dy + hy * dh, hw * dw, hh * dh
        C.highlight(slide, bx, by, bw, bh)
        if isinstance(hl, dict) and hl.get("label"):
            # label inside the figure box: below the frame if there is room, else above it
            lw, lh = max(bw, 3.0), 0.5
            lx = min(max(bx, x), x + w - lw)
            ly = by + bh + 0.05 if by + bh + 0.05 + lh <= y + h else max(by - lh - 0.05, y)
            C.textbox(slide, lx, ly, lw, lh, hl["label"], style="key", size=18)
    return pic


def draw_shapes(slide, layout_name, shapes, base):
    for s in shapes:
        t = s["type"]
        frame = s.get("frame", "abs")
        if "box" in s:
            x, y, w, h = zone_box(layout_name, s["box"], frame)
        if t == "box":
            C.box(slide, x, y, w, h, s.get("text", ""), kind=s.get("kind", "default"),
                  size=s.get("size", 24), align=s.get("align", "c"), shape=s.get("shape", "round"))
        elif t == "text":
            C.textbox(slide, x, y, w, h, s["text"], style=s.get("style", "body"),
                      size=s.get("size", 24), align=s.get("align", "l"), anchor=s.get("anchor", "t"),
                      bullets=s.get("bullets", False))
        elif t == "arrow":
            (x1, y1), (x2, y2) = s["from"], s["to"]
            if frame == "zone":
                x1, y1, _, _ = zone_box(layout_name, (x1, y1, 0, 0), "zone")
                x2, y2, _, _ = zone_box(layout_name, (x2, y2, 0, 0), "zone")
            C.arrow(slide, x1, y1, x2, y2, kind=s.get("kind", "default"), head=s.get("head", "end"))
        elif t == "highlight":
            C.highlight(slide, x, y, w, h)
        elif t == "flow":
            C.flow(slide, x, y, w, h, s["items"], key=set(s.get("key", [])),
                   emph=set(s.get("emph", [])), size=s.get("size", 24),
                   vertical=s.get("vertical", False), gap=s.get("gap", 0.55))
        elif t == "chevrons":
            C.chevrons(slide, x, y, w, h, s["items"], key=set(s.get("key", [])), size=s.get("size", 22))
        elif t == "image":
            C.image_fit(slide, resolve(base, s["path"]), x, y, w, h, crop=s.get("crop"),
                        align=s.get("align", "c"))
        elif t == "math":
            if s.get("render") == "image":  # only when an editable equation is not wanted
                png = C.math_image(s["tex"], size=s.get("size", 40))
                C.image_fit(slide, png, x, y, w, h, align=s.get("align", "c"))
            else:
                C.equation(slide, x, y, w, h, s["tex"], size=s.get("size", 36),
                           align=s.get("align", "c"))
        elif t == "table":
            C.table(slide, x, y, w, s["rows"], col_widths=s.get("col_widths"),
                    size=s.get("size", 20), row_h=s.get("row_h", 0.6),
                    ours_rows=set(s.get("ours_rows", [])), key_cells=s.get("key_cells", []),
                    first_col=s.get("first_col", True), align=s.get("align", "c"))
        elif t == "chart":
            C.chart(slide, x, y, w, h, s["categories"], s["series"], kind=s.get("kind", "bar"),
                    ylabel=s.get("ylabel"), legend=s.get("legend", True), labels=s.get("labels", False),
                    size=s.get("size", 18), number_format=s.get("number_format"))
        elif t == "swatch":
            C.swatch(slide, x, y, w, h, s["color"], s.get("label", ""), s.get("sub", ""))
        else:
            raise SystemExit(f"unknown shape type '{t}'")


TITLE_MAX_WORDS = 4
_SKIP = {"vs", "vs.", "·", "-", "–", "—", "&", "/", ":"}


def lint_title(title):
    """Lab title style: noun phrase, at most 4 words, noun-form ending (no sentence)."""
    if not title or not isinstance(title, str):
        return None
    words = [w for w in re.split(r"\s+", title.strip()) if w and w not in _SKIP]
    probs = []
    if len(words) > TITLE_MAX_WORDS:
        probs.append(f"{len(words)} words (max {TITLE_MAX_WORDS})")
    if re.search(r"(다|[어아해세이]요|죠|까|[.!?])$", title.strip()):
        probs.append("sentence ending; use a noun phrase")
    return f"title '{title}': " + ", ".join(probs) if probs else None


def add_slide(prs, sd, meta, base):
    layout = find_layout(prs, sd["layout"])
    slide = prs.slides.add_slide(layout)
    roles = layout_roles(layout)
    by_idx = {p.placeholder_format.idx: p for p in slide.placeholders}

    if sd["layout"] in CONTENT_LAYOUTS:
        w_ = lint_title(sd.get("title"))
        if w_:
            WARNINGS.append(w_)
    values = dict(sd)
    if sd["layout"] in CONTENT_LAYOUTS and "source" not in values and meta.get("source"):
        values["source"] = meta["source"]
    if sd["layout"] in ("Title", "Closing") and "date" not in values and meta.get("date"):
        values["date"] = meta["date"]

    used = set()
    for role, idx in roles.items():
        if role not in values or values[role] in (None, "", []):
            continue
        ph = by_idx.get(idx)
        if ph is None:
            continue
        val = values[role]
        used.add(idx)
        if ph.placeholder_format.type is not None and "PICTURE" in str(ph.placeholder_format.type):
            place_figure(slide, ph, val, base)
        elif role == "toc":
            items = val["items"] if isinstance(val, dict) else val
            cur = val.get("current") if isinstance(val, dict) else None
            tf = ph.text_frame
            tf.clear()
            for i, it in enumerate(items):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                C.add_runs(p, f"{i+1:02d}\t{it}")
                if cur is not None and i == cur:
                    for r in p.runs:
                        C.set_color(r.font.color, "accent1")
                        r.font.bold = True
        else:
            C.fill_text(ph.text_frame, val, levels=role in TEXT_ROLES_WITH_LEVELS)

    # drop placeholders that were not filled
    for idx, ph in by_idx.items():
        if idx not in used and ph._element.getparent() is not None:
            ph._element.getparent().remove(ph._element)

    if sd.get("shapes"):
        draw_shapes(slide, sd["layout"], sd["shapes"], base)
    import rs_math
    rs_math.wrap_inline_math(slide)  # after all text is written
    if sd.get("notes"):
        slide.notes_slide.notes_text_frame.text = sd["notes"]
    if sd.get("hidden"):
        slide._element.set("show", "0")
    return slide


def build(spec_path, out, template=None, keep=False):
    spec_path = Path(spec_path)
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    base = spec_path.parent
    meta = spec.get("meta", {})
    tpl = Path(template or spec.get("template") or DEFAULT_TEMPLATE)
    if not tpl.is_absolute() and not tpl.exists():
        tpl = resolve(base, tpl)
    prs = Presentation(tpl)
    if not keep:
        clear_slides(prs)
    if meta.get("footer"):
        set_footer(prs, meta["footer"])
    for i, sd in enumerate(spec["slides"], 1):
        n_warn = len(WARNINGS)
        try:
            add_slide(prs, sd, meta, base)
        except Exception as e:  # make spec errors easy to locate
            raise SystemExit(f"slide {i} ({sd.get('layout')}): {e}")
        for k in range(n_warn, len(WARNINGS)):
            WARNINGS[k] = f"slide {i}: {WARNINGS[k]}"
    if meta.get("title"):
        prs.core_properties.title = meta["title"]
    prs.save(out)
    print(f"wrote {out}: {len(prs.slides)} slides")
    for w_ in WARNINGS:
        print("WARNING", w_)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("out")
    ap.add_argument("--template")
    ap.add_argument("--keep-slides", action="store_true", help="keep slides already in the template")
    a = ap.parse_args()
    build(a.spec, a.out, a.template, a.keep_slides)
