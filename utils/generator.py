"""
Batch certificate generation and ZIP archiving engine for Bulk Certificate Generator.
"""
from io import BytesIO
import zipfile
import re
from typing import Dict, Any, List, Optional, Callable, Tuple
import pandas as pd
from PIL import Image, ImageDraw

from utils.font_manager import get_pillow_font
import pypdf


def sanitize_filename(name: str) -> str:
    """
    Sanitizes string for safe use in file and archive paths.
    """
    clean = re.sub(r'[\\/*?:"<>|]', '_', str(name).strip())
    # Collapse multiple underscores or spaces
    clean = re.sub(r'\s+', ' ', clean)
    return clean or "certificate"


def format_certificate_filename(
    pattern: str,
    name: str,
    roll: str,
    idx: int
) -> str:
    """
    Formats the PDF filename according to selected naming convention.
    """
    safe_name = sanitize_filename(name)
    safe_roll = sanitize_filename(roll)

    if pattern == "StudentName_RollNumber.pdf":
        base = f"{safe_name}_{safe_roll}"
    elif pattern == "StudentName.pdf":
        base = safe_name
    elif pattern == "Certificate_RollNumber.pdf":
        base = f"Certificate_{safe_roll}"
    else:  # Default: RollNumber_StudentName.pdf
        base = f"{safe_roll}_{safe_name}"

    return f"{base}.pdf"


def render_single_certificate(
    base_template: Image.Image,
    student_record: Dict[str, Any],
    name_col: str,
    roll_col: str,
    extra_col: Optional[str],
    fields_config: Dict[str, Any]
) -> Image.Image:
    """
    Renders personalized text onto a fresh copy of the certificate template.
    """
    cert = base_template.copy()
    draw = ImageDraw.Draw(cert)

    # 1. Student Name
    raw_name = str(student_record.get(name_col, "")).strip()
    if raw_name and "name" in fields_config:
        cfg = fields_config["name"]
        font = get_pillow_font(cfg.get("font_family", "Arial"), cfg.get("font_style", "Bold"), cfg.get("font_size", 20))
        anchor = "mm" if cfg.get("align", "center") == "center" else ("lm" if cfg.get("align", "center") == "left" else "rm")
        draw.text((cfg["x"], cfg["y"]), raw_name, fill=cfg.get("color", "#1A2B4C"), font=font, anchor=anchor)

    # 2. Roll Number
    raw_roll = str(student_record.get(roll_col, "")).strip()
    if raw_roll and "roll" in fields_config:
        cfg = fields_config["roll"]
        prefix = cfg.get("prefix", "")
        roll_text = f"{prefix}{raw_roll}"
        font = get_pillow_font(cfg.get("font_family", "Arial"), cfg.get("font_style", "Bold"), cfg.get("font_size", 11))
        anchor = "mm" if cfg.get("align", "center") == "center" else ("lm" if cfg.get("align", "center") == "left" else "rm")
        draw.text((cfg["x"], cfg["y"]), roll_text, fill=cfg.get("color", "#1A2B4C"), font=font, anchor=anchor)

    # 3. Extra Field
    if extra_col and extra_col != "(None)" and extra_col in student_record and "extra" in fields_config:
        raw_extra = str(student_record.get(extra_col, "")).strip()
        if raw_extra:
            cfg = fields_config["extra"]
            prefix = cfg.get("prefix", "")
            extra_text = f"{prefix}{raw_extra}"
            font = get_pillow_font(cfg.get("font_family", "Arial"), cfg.get("font_style", "Regular"), cfg.get("font_size", 14))
            anchor = "mm" if cfg.get("align", "center") == "center" else ("lm" if cfg.get("align", "center") == "left" else "rm")
            draw.text((cfg["x"], cfg["y"]), extra_text, fill=cfg.get("color", "#475569"), font=font, anchor=anchor)

    return cert


def generate_combined_pdf(pdf_bytes_list: List[bytes]) -> bytes:
    """
    Combines a list of individual PDF bytes into one single multi-page PDF document,
    maintaining the exact original order. Each student has exactly one page.
    """
    writer = pypdf.PdfWriter()
    for pdf_data in pdf_bytes_list:
        writer.append(BytesIO(pdf_data))
    out_buf = BytesIO()
    writer.write(out_buf)
    return out_buf.getvalue()


def generate_batch_zip(
    base_template: Image.Image,
    valid_df: pd.DataFrame,
    name_col: str,
    roll_col: str,
    extra_col: Optional[str],
    fields_config: Dict[str, Any],
    naming_pattern: str = "RollNumber_StudentName.pdf",
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> Tuple[bytes, bytes, List[Dict[str, Any]]]:
    """
    Renders high-resolution individual PDF certificates for all valid records,
    combines them into a single multi-page PDF (All_Student_Certificates.pdf),
    and packages individual PDFs into an in-memory ZIP archive.

    Returns: (zip_bytes, combined_pdf_bytes, generated_file_manifest)
    """
    zip_buffer = BytesIO()
    manifest = []
    pdf_bytes_list = []
    total_records = len(valid_df)

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (_, row) in enumerate(valid_df.iterrows(), start=1):
            record = dict(row)
            student_name = str(record.get(name_col, f"Student_{idx}")).strip()
            student_roll = str(record.get(roll_col, f"ID_{idx}")).strip()

            if progress_callback:
                progress_callback(idx, total_records, student_name)

            # Render certificate image with personalized data
            cert_img = render_single_certificate(
                base_template=base_template,
                student_record=record,
                name_col=name_col,
                roll_col=roll_col,
                extra_col=extra_col,
                fields_config=fields_config
            )

            # Convert to PDF in memory
            pdf_buf = BytesIO()
            cert_img.save(pdf_buf, format="PDF", resolution=150.0)
            pdf_data = pdf_buf.getvalue()
            pdf_bytes_list.append(pdf_data)

            # Formulate safe filename
            pdf_filename = format_certificate_filename(
                pattern=naming_pattern,
                name=student_name,
                roll=student_roll,
                idx=idx
            )

            # Handle possible collision in filename
            unique_name = pdf_filename
            counter = 1
            while unique_name in [m["filename"] for m in manifest]:
                unique_name = f"{pdf_filename[:-4]}_{counter}.pdf"
                counter += 1

            # Write into ZIP archive
            zf.writestr(unique_name, pdf_data)

            manifest.append({
                "index": idx,
                "filename": unique_name,
                "name": student_name,
                "roll": student_roll,
                "size_bytes": len(pdf_data)
            })

    # Combine all individual PDFs into a single multi-page PDF document
    combined_pdf_bytes = generate_combined_pdf(pdf_bytes_list)

    return zip_buffer.getvalue(), combined_pdf_bytes, manifest
