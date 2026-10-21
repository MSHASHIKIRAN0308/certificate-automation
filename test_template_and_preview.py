"""
Unit test suite verifying PDF, PNG, JPG template loading, portrait/landscape orientation,
Telugu and English font typography, and text overlap/boundary collision detection.
"""
import os
from io import BytesIO
from PIL import Image

from utils.preview_renderer import load_template_image, render_preview_overlay
from utils.font_manager import get_pillow_font, AVAILABLE_FONT_FAMILIES


class DummyFile:
    def __init__(self, path: str):
        self.name = os.path.basename(path)
        with open(path, "rb") as f:
            self.data = f.read()

    def getvalue(self):
        return self.data


def test_template_formats():
    # 1. Test PDF Template Loading
    pdf_path = os.path.join("sample_data", "sample_certificate_template.pdf")
    img_pdf, meta_pdf, err = load_template_image(DummyFile(pdf_path))
    assert err is None, f"PDF load error: {err}"
    assert meta_pdf["is_pdf"] is True
    assert meta_pdf["format"] == "PDF"
    assert meta_pdf["orientation"] == "Landscape"
    assert img_pdf.size == (meta_pdf["width"], meta_pdf["height"])
    print("[PASS] PDF template loaded with metadata:", meta_pdf["width"], "x", meta_pdf["height"])

    # 2. Test JPG Template Loading
    jpg_path = os.path.join("sample_data", "sreenidhi_certificate_template.jpg")
    img_jpg, meta_jpg, err = load_template_image(DummyFile(jpg_path))
    assert err is None, f"JPG load error: {err}"
    assert meta_jpg["format"] == "JPEG"
    assert meta_jpg["width"] == 717
    assert meta_jpg["height"] == 505
    print("[PASS] JPG template loaded successfully!")

    # 3. Test Portrait PNG Template Loading
    portrait_path = os.path.join("sample_data", "sample_portrait_template.png")
    img_port, meta_port, err = load_template_image(DummyFile(portrait_path))
    assert err is None, f"Portrait load error: {err}"
    assert meta_port["orientation"] == "Portrait"
    assert meta_port["width"] < meta_port["height"]
    print("[PASS] Portrait template identified:", meta_port["orientation"])


def test_telugu_and_english_fonts():
    # Verify Telugu font loading
    font_telugu = get_pillow_font("Noto Sans Telugu", "Bold", 24)
    assert font_telugu is not None

    font_english = get_pillow_font("Arial", "Bold", 24)
    assert font_english is not None
    print("[PASS] Telugu and English fonts loaded successfully!")


def test_overlap_and_boundary_detection():
    base_img = Image.new("RGB", (600, 400), color="white")

    # Overlapping fields test (same Y position)
    fields_overlapping = {
        "name": {
            "text": "Sophia Montgomery",
            "x": 300,
            "y": 200,
            "font_size": 24,
            "font_family": "Arial",
            "font_style": "Bold",
            "align": "center",
            "label": "Student Name"
        },
        "roll": {
            "text": "CS2026-001",
            "x": 300,
            "y": 205,  # Very close to name -> collision!
            "font_size": 20,
            "font_family": "Arial",
            "font_style": "Regular",
            "align": "center",
            "label": "Roll Number"
        }
    }

    _, warnings, _ = render_preview_overlay(base_img, fields_overlapping)
    assert len(warnings) > 0, "Expected overlap warning, got none"
    assert any("Text Collision Detected" in w for w in warnings)
    print("[PASS] Overlap collision correctly detected!")

    # Non-overlapping fields test
    fields_clean = {
        "name": {
            "text": "రాహుల్ వర్మ (Rahul Varma)",
            "x": 300,
            "y": 150,
            "font_size": 22,
            "font_family": "Noto Sans Telugu",
            "font_style": "Bold",
            "align": "center",
            "label": "Student Name"
        },
        "roll": {
            "text": "21A91A0501",
            "x": 300,
            "y": 250,
            "font_size": 16,
            "font_family": "Arial",
            "font_style": "Regular",
            "align": "center",
            "label": "Roll Number"
        }
    }

    _, clean_warnings, bboxes = render_preview_overlay(base_img, fields_clean)
    assert len(clean_warnings) == 0, f"Expected no warnings, got: {clean_warnings}"
    assert "name" in bboxes and "roll" in bboxes
    print("[PASS] Clean non-overlapping layout verified without false alarms!")


if __name__ == "__main__":
    test_template_formats()
    test_telugu_and_english_fonts()
    test_overlap_and_boundary_detection()
    print("[SUCCESS] ALL TEMPLATE AND PREVIEW TESTS PASSED!")
