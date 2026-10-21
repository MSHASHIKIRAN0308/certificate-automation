"""
Preview and template rendering utilities for Bulk Certificate Generator.
Supports PDF, PNG, JPG templates with multi-language typography (Telugu & English)
and intelligent text overlap/boundary collision detection.
"""
from io import BytesIO
from typing import Tuple, Optional, Dict, Any, List
from PIL import Image, ImageDraw
import pymupdf

from utils.font_manager import get_pillow_font


def load_template_image(uploaded_file) -> Tuple[Optional[Image.Image], Optional[Dict[str, Any]], Optional[str]]:
    """
    Loads uploaded template image (PNG, JPG, JPEG) or the first page of a PDF template into a PIL Image.
    Preserves original dimensions, color depth, vector resolution and design integrity.
    Returns: (PIL.Image, metadata_dict, error_message)
    """
    if uploaded_file is None:
        return None, None, "No template file provided."

    try:
        filename = uploaded_file.name.lower()
        file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

        is_pdf = filename.endswith(".pdf")
        page_count = 1

        if is_pdf:
            doc = pymupdf.open(stream=file_bytes, filetype="pdf")
            if len(doc) == 0:
                return None, None, "The uploaded PDF document contains no pages."
            page_count = len(doc)
            page = doc.load_page(0)
            
            # High-resolution 2.0x matrix render for sharp vector preview
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2.0, 2.0), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            template_format = "PDF"
        elif filename.endswith((".png", ".webp")):
            img = Image.open(BytesIO(file_bytes)).convert("RGB")
            template_format = "PNG"
        elif filename.endswith((".jpg", ".jpeg")):
            img = Image.open(BytesIO(file_bytes)).convert("RGB")
            template_format = "JPEG"
        else:
            return None, None, "Unsupported template format. Please upload a PDF, PNG, or JPG file."

        w, h = img.size
        orientation = "Landscape" if w >= h else "Portrait"
        aspect_ratio = round(w / h, 2)

        metadata = {
            "is_pdf": is_pdf,
            "format": template_format,
            "width": w,
            "height": h,
            "orientation": orientation,
            "aspect_ratio": aspect_ratio,
            "page_count": page_count,
            "raw_bytes": file_bytes,
            "filename": uploaded_file.name
        }

        return img, metadata, None
    except Exception as e:
        return None, None, f"Failed to load certificate template: {str(e)}"


def create_default_template_image(width: int = 1920, height: int = 1080) -> Image.Image:
    """
    Generates an elegant, high-resolution default certificate background template for testing.
    """
    img = Image.new("RGB", (width, height), color="#0F172A")
    draw = ImageDraw.Draw(img)

    # Outer ornate borders
    draw.rectangle([40, 40, width - 40, height - 40], outline="#D97706", width=4)
    draw.rectangle([55, 55, width - 55, height - 55], outline="#1E293B", width=2)
    draw.rectangle([70, 70, width - 70, height - 70], outline="#F59E0B", width=1)

    # Corner decorations
    corner_size = 60
    corners = [
        (40, 40, 40 + corner_size, 40 + corner_size),
        (width - 40 - corner_size, 40, width - 40, 40 + corner_size),
        (40, height - 40 - corner_size, 40 + corner_size, height - 40),
        (width - 40 - corner_size, height - 40 - corner_size, width - 40, height - 40),
    ]
    for c in corners:
        draw.rectangle(c, outline="#D97706", width=2)

    font_large = get_pillow_font("Georgia", "Bold", 58)
    font_sub = get_pillow_font("Arial", "Regular", 26)
    font_small = get_pillow_font("Arial", "Regular", 20)

    # Title & Sub-header
    draw.text((width // 2, 160), "CERTIFICATE OF EXCELLENCE", fill="#F8FAFC", font=font_large, anchor="mm")
    draw.text((width // 2, 230), "THIS IS PROUDLY PRESENTED TO", fill="#94A3B8", font=font_sub, anchor="mm")

    # Lower context text
    draw.text((width // 2, 630), "FOR OUTSTANDING PERFORMANCE AND ACTIVE PARTICIPATION", fill="#94A3B8", font=font_sub, anchor="mm")
    draw.text((width // 2, 680), "ACADEMIC EXCELLENCE & SKILL DEVELOPMENT PROGRAM", fill="#CBD5E1", font=font_small, anchor="mm")

    # Signature line placeholders
    sig_y = height - 180
    draw.line([(250, sig_y), (550, sig_y)], fill="#64748B", width=2)
    draw.text((400, sig_y + 25), "PROGRAM DIRECTOR", fill="#94A3B8", font=font_small, anchor="mm")

    draw.line([(width - 550, sig_y), (width - 250, sig_y)], fill="#64748B", width=2)
    draw.text((width - 400, sig_y + 25), "AUTHORIZED SIGNATORY", fill="#94A3B8", font=font_small, anchor="mm")

    return img


def check_bbox_intersection(box_a: Tuple[int, int, int, int], box_b: Tuple[int, int, int, int]) -> bool:
    """
    Returns True if two bounding boxes (x0, y0, x1, y1) intersect/overlap.
    """
    return not (
        box_a[2] < box_b[0] or  # a is completely to left of b
        box_a[0] > box_b[2] or  # a is completely to right of b
        box_a[3] < box_b[1] or  # a is completely above b
        box_a[1] > box_b[3]     # a is completely below b
    )


def render_preview_overlay(
    base_image: Image.Image,
    fields: Dict[str, Dict[str, Any]],
    show_bounding_box: bool = True,
    show_anchors: bool = True
) -> Tuple[Image.Image, List[str], Dict[str, Any]]:
    """
    Renders preview text overlay on top of the base template image.
    Calculates exact bounding boxes, detects overlaps between fields,
    and checks for text overflow outside the certificate canvas.

    Returns: (annotated_image, list_of_warnings, field_bboxes)
    """
    preview = base_image.copy()
    draw = ImageDraw.Draw(preview)
    img_w, img_h = preview.size

    warnings = []
    field_bboxes = {}

    # First pass: compute fonts and bounding boxes for all active fields
    computed_fields = {}
    for field_key, config in fields.items():
        text = str(config.get("text", "")).strip()
        if not text:
            continue

        x = int(config.get("x", img_w // 2))
        y = int(config.get("y", img_h // 2))
        size = int(config.get("font_size", 32))
        color = config.get("color", "#000000")
        align = config.get("align", "center")
        font_family = config.get("font_family", "Arial")
        font_style = config.get("font_style", "Bold")

        font = get_pillow_font(font_family, font_style, size)

        # Anchor mapping
        if align == "center":
            anchor = "mm"
        elif align == "left":
            anchor = "lm"
        else:
            anchor = "rm"

        bbox = draw.textbbox((x, y), text, font=font, anchor=anchor)
        field_bboxes[field_key] = bbox

        computed_fields[field_key] = {
            "text": text,
            "x": x,
            "y": y,
            "font": font,
            "color": color,
            "anchor": anchor,
            "bbox": bbox,
            "label": config.get("label", field_key)
        }

        # Check for canvas boundary overflow
        overflow_x = bbox[0] < 0 or bbox[2] > img_w
        overflow_y = bbox[1] < 0 or bbox[3] > img_h
        if overflow_x or overflow_y:
            label = config.get("label", field_key)
            warnings.append(
                f"⚠️ **{label}** exceeds template canvas boundaries! "
                f"(Text box: [{bbox[0]}, {bbox[1]}, {bbox[2]}, {bbox[3]}], Canvas: {img_w}x{img_h}). "
                f"Adjust position or decrease font size."
            )

    # Second pass: detect overlaps between fields
    keys = list(computed_fields.keys())
    for i in range(len(keys)):
        for j in range(i + 1, len(keys)):
            k1, k2 = keys[i], keys[j]
            box1 = computed_fields[k1]["bbox"]
            box2 = computed_fields[k2]["bbox"]

            if check_bbox_intersection(box1, box2):
                lbl1 = computed_fields[k1]["label"]
                lbl2 = computed_fields[k2]["label"]
                warnings.append(
                    f"⚠️ **Text Collision Detected:** '{lbl1}' overlaps with '{lbl2}'! "
                    f"Please adjust the vertical Y position or decrease font size to keep text separated."
                )

    # Third pass: render text onto preview image
    for field_key, item in computed_fields.items():
        draw.text(
            (item["x"], item["y"]),
            item["text"],
            fill=item["color"],
            font=item["font"],
            anchor=item["anchor"]
        )

        # Optional visual bounding box guide
        if show_bounding_box:
            b = item["bbox"]
            # Draw semi-transparent dotted/thin rectangle around text
            draw.rectangle([b[0] - 2, b[1] - 2, b[2] + 2, b[3] + 2], outline="#0284C7", width=1)

        # Optional anchor crosshair
        if show_anchors:
            ax, ay = item["x"], item["y"]
            arm = 8
            draw.line([(ax - arm, ay), (ax + arm, ay)], fill="#E11D48", width=2)
            draw.line([(ax, ay - arm), (ax, ay + arm)], fill="#E11D48", width=2)

    return preview, warnings, field_bboxes
