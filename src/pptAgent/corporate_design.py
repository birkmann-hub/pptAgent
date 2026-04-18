"""Corporate design constants and helper utilities."""

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor

# ---------------------------------------------------------------------------
# Corporate colour palette
# ---------------------------------------------------------------------------

# Primary
DARK_BLUE = RGBColor(0x00, 0x33, 0x66)        # #003366
MEDIUM_BLUE = RGBColor(0x00, 0x66, 0xCC)       # #0066CC
LIGHT_BLUE = RGBColor(0x66, 0xAA, 0xDD)        # #66AADD

# Accent
ACCENT_ORANGE = RGBColor(0xFF, 0x66, 0x00)     # #FF6600
ACCENT_TEAL = RGBColor(0x00, 0x99, 0x99)       # #009999

# Neutrals
WHITE = RGBColor(0xFF, 0xFF, 0xFF)             # #FFFFFF
LIGHT_GRAY = RGBColor(0xF0, 0xF2, 0xF5)       # #F0F2F5
MID_GRAY = RGBColor(0xAA, 0xAA, 0xAA)         # #AAAAAA
DARK_TEXT = RGBColor(0x1A, 0x1A, 0x2E)        # #1A1A2E
GRAY = RGBColor(0x77, 0x77, 0x77)             # #777777

# Colour name → RGBColor mapping used when parsing YAML definitions
COLOUR_MAP: dict[str, RGBColor] = {
    "dark_blue": DARK_BLUE,
    "medium_blue": MEDIUM_BLUE,
    "light_blue": LIGHT_BLUE,
    "accent_orange": ACCENT_ORANGE,
    "accent_teal": ACCENT_TEAL,
    "white": WHITE,
    "light_gray": LIGHT_GRAY,
    "mid_gray": MID_GRAY,
    "dark_text": DARK_TEXT,
    "gray": GRAY,
}

# ---------------------------------------------------------------------------
# Typography
# ---------------------------------------------------------------------------

FONT_FAMILY = "Calibri"
FONT_FAMILY_HEADINGS = "Calibri"

# ---------------------------------------------------------------------------
# Slide dimensions (standard widescreen 16:9)
# ---------------------------------------------------------------------------

SLIDE_WIDTH_EMU = 9_144_000   # 25.4 cm
SLIDE_HEIGHT_EMU = 5_143_500  # 14.29 cm

# ---------------------------------------------------------------------------
# Layout zones (as fractions of slide dimensions)
# ---------------------------------------------------------------------------

HEADER_HEIGHT_FRAC = 0.14   # top stripe containing slide title
FOOTER_HEIGHT_FRAC = 0.08   # bottom stripe
SIDE_MARGIN_FRAC = 0.04     # left / right outer margin
INNER_MARGIN_FRAC = 0.02    # gap between columns / elements

# ---------------------------------------------------------------------------
# Derived pixel-equivalent positions (in EMU)
# ---------------------------------------------------------------------------

def _emu(fraction: float, total: int) -> Emu:
    return Emu(int(fraction * total))

HEADER_TOP = Emu(0)
HEADER_HEIGHT = _emu(HEADER_HEIGHT_FRAC, SLIDE_HEIGHT_EMU)
CONTENT_TOP = _emu(HEADER_HEIGHT_FRAC, SLIDE_HEIGHT_EMU)
CONTENT_HEIGHT = _emu(
    1 - HEADER_HEIGHT_FRAC - FOOTER_HEIGHT_FRAC, SLIDE_HEIGHT_EMU
)
FOOTER_TOP = _emu(1 - FOOTER_HEIGHT_FRAC, SLIDE_HEIGHT_EMU)
FOOTER_HEIGHT = _emu(FOOTER_HEIGHT_FRAC, SLIDE_HEIGHT_EMU)
LEFT_MARGIN = _emu(SIDE_MARGIN_FRAC, SLIDE_WIDTH_EMU)
CONTENT_WIDTH = _emu(1 - 2 * SIDE_MARGIN_FRAC, SLIDE_WIDTH_EMU)

# ---------------------------------------------------------------------------
# Helper: resolve a colour alias or hex string to RGBColor
# ---------------------------------------------------------------------------

def resolve_color(value: str | RGBColor | None) -> RGBColor:
    """Return an RGBColor for an alias name (e.g. 'dark_blue') or hex string."""
    if value is None:
        return DARK_TEXT
    if isinstance(value, RGBColor):
        return value
    name = str(value).lower().strip()
    if name in COLOUR_MAP:
        return COLOUR_MAP[name]
    # Try as hex string (#RRGGBB or RRGGBB)
    hex_val = name.lstrip("#")
    if len(hex_val) == 6:
        try:
            r = int(hex_val[0:2], 16)
            g = int(hex_val[2:4], 16)
            b = int(hex_val[4:6], 16)
            return RGBColor(r, g, b)
        except ValueError:
            pass
    return DARK_TEXT
