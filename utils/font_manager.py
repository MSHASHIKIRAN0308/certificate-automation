"""
Font Manager utility supporting English and Telugu Unicode typography for Bulk Certificate Generator.
"""
import os
from typing import Dict, List, Optional
from PIL import ImageFont


AVAILABLE_FONT_FAMILIES = [
    "Arial",
    "Calibri",
    "Times New Roman",
    "Georgia",
    "Noto Sans Telugu",
    "Nirmala UI"
]

FONT_STYLES = ["Regular", "Bold", "Italic", "Bold Italic"]


def _find_font_file(font_family: str, style: str) -> Optional[str]:
    """
    Locates the TTF/TTC file for a given font family and weight style.
    Looks in local ./fonts directory first, then system Windows font folder.
    """
    style_lower = style.lower()
    is_bold = "bold" in style_lower
    is_italic = "italic" in style_lower

    local_fonts_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "fonts")
    win_fonts_dir = "C:/Windows/Fonts"

    # 1. Telugu Noto Sans
    if "telugu" in font_family.lower() or font_family == "Noto Sans Telugu":
        if is_bold:
            candidate = os.path.join(local_fonts_dir, "NotoSansTelugu-Bold.ttf")
            if os.path.exists(candidate):
                return candidate
        candidate = os.path.join(local_fonts_dir, "NotoSansTelugu-Regular.ttf")
        if os.path.exists(candidate):
            return candidate

        # System fallback for Telugu: Nirmala.ttc
        nirmala = os.path.join(win_fonts_dir, "Nirmala.ttc")
        if os.path.exists(nirmala):
            return nirmala

    # 2. Nirmala UI (Windows Telugu & Indic Unicode)
    if "nirmala" in font_family.lower():
        nirmala = os.path.join(win_fonts_dir, "Nirmala.ttc")
        if os.path.exists(nirmala):
            return nirmala

    # 3. Standard Windows Fonts
    font_map = {
        "arial": {
            "regular": "arial.ttf",
            "bold": "arialbd.ttf",
            "italic": "ariali.ttf",
            "bold italic": "arialbi.ttf"
        },
        "calibri": {
            "regular": "calibri.ttf",
            "bold": "calibrib.ttf",
            "italic": "calibrii.ttf",
            "bold italic": "calibriz.ttf"
        },
        "times new roman": {
            "regular": "times.ttf",
            "bold": "timesbd.ttf",
            "italic": "timesi.ttf",
            "bold italic": "timesbi.ttf"
        },
        "georgia": {
            "regular": "georgia.ttf",
            "bold": "georgiab.ttf",
            "italic": "georgiai.ttf",
            "bold italic": "georgiaz.ttf"
        }
    }

    fam_key = font_family.lower().strip()
    if fam_key in font_map:
        style_key = "regular"
        if is_bold and is_italic:
            style_key = "bold italic"
        elif is_bold:
            style_key = "bold"
        elif is_italic:
            style_key = "italic"

        filename = font_map[fam_key].get(style_key, font_map[fam_key]["regular"])
        win_path = os.path.join(win_fonts_dir, filename)
        if os.path.exists(win_path):
            return win_path

    # Fallback search
    for test_dir in [local_fonts_dir, win_fonts_dir]:
        if os.path.exists(test_dir):
            for f in os.listdir(test_dir):
                if fam_key in f.lower() and f.lower().endswith((".ttf", ".ttc")):
                    return os.path.join(test_dir, f)

    return None


def get_pillow_font(font_family: str, style: str, size: int) -> ImageFont.ImageFont:
    """
    Returns a PIL ImageFont object configured for the requested family, style, and size.
    """
    font_path = _find_font_file(font_family, style)
    if font_path:
        try:
            return ImageFont.truetype(font_path, size)
        except Exception:
            pass

    # Fallback to system Arial or default
    try:
        return ImageFont.truetype("arial.ttf", size)
    except Exception:
        return ImageFont.load_default()
