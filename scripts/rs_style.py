"""research-slides: shared design tokens and small helpers.

Single source of truth for palette, fonts, and geometry. Both make_template.py
and build_deck.py import this module, so a change here changes everything.
"""

# ---------------------------------------------------------------------------
# Palette A (Pantone TCX, sRGB approximations)
# Theme slot -> (hex, Pantone name, code, role)
# ---------------------------------------------------------------------------
PALETTE = {
    "dk1":     ("262626", "-",             "-",       "본문 텍스트"),
    "lt1":     ("FFFFFF", "-",             "-",       "배경"),
    "dk2":     ("223A5E", "Navy Peony",    "19-4029", "구조 (dk2 = accent1)"),
    "lt2":     ("F3F4F6", "-",             "-",       "옅은 회색 채움"),
    "accent1": ("223A5E", "Navy Peony",    "19-4029", "구조: 제목, 기본 도형, 표 헤더, 바"),
    "accent2": ("AD5D5D", "Dusty Cedar",   "18-1630", "핵심 강조: 핵심어, 강조 박스, Ours"),
    "accent3": ("658DC6", "Provence",      "16-4032", "보조: 두 번째 계열, 보조 선"),
    "accent4": ("7BC4C4", "Aqua Sky",      "14-4811", "검증·정답: 정답 표시, 통과"),
    "accent5": ("F0C05A", "Mimosa",        "14-0848", "낮은 우선순위: 참고, 부가 정보"),
    "accent6": ("939597", "Ultimate Gray", "17-5104", "중립: baseline, 화살표, 비활성"),
    "hlink":   ("658DC6", "Provence",      "16-4032", "링크"),
    "folHlink": ("939597", "Ultimate Gray", "17-5104", "방문한 링크"),
}

HEX = {k: v[0] for k, v in PALETTE.items()}

FONT_REGULAR = "NanumSquareOTF"
FONT_HEAVY = "NanumSquareOTF ExtraBold"

# Slide geometry (inches). 16:9 at 20 x 11.25 in, same as the lab's decks.
SLIDE_W = 20.0
SLIDE_H = 11.25
EMU = 914400


def emu(inches: float) -> int:
    return int(round(inches * EMU))


# ---------------------------------------------------------------------------
# Colour spec mini-language used everywhere in the skill:
#   "accent1"      -> theme colour
#   "accent1:12"   -> 12% of the colour mixed with white (a tint, for fills)
#   "tx1:75"       -> 75% of the colour mixed with white (for grey text)
#   "#AD5D5D"      -> literal hex (avoid; only for imported brand colours)
# ---------------------------------------------------------------------------
def clr_xml(spec: str) -> str:
    """Return DrawingML colour element XML for a colour spec."""
    if spec.startswith("#"):
        return f'<a:srgbClr val="{spec[1:].upper()}"/>'
    if ":" in spec:
        name, pct = spec.split(":")
        pct = int(pct)
        return (f'<a:schemeClr val="{name}"><a:lumMod val="{pct*1000}"/>'
                f'<a:lumOff val="{(100-pct)*1000}"/></a:schemeClr>')
    return f'<a:schemeClr val="{spec}"/>'


def fill_xml(spec):
    if spec is None:
        return "<a:noFill/>"
    return f"<a:solidFill>{clr_xml(spec)}</a:solidFill>"


def line_xml(spec, width_pt=1.0, dash=None):
    if spec is None:
        return "<a:ln><a:noFill/></a:ln>"
    d = f'<a:prstDash val="{dash}"/>' if dash else ""
    return (f'<a:ln w="{int(width_pt*12700)}">{fill_xml(spec)}{d}</a:ln>')
