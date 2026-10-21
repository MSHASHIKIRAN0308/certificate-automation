"""
Bulk Certificate Generator - Streamlit Web Application
A secure, 100% local tool to generate individual PDF certificates from Excel & CSV records.
Supports PDF, PNG, JPG templates in Landscape & Portrait orientations with Telugu & English typography.
"""
import os
import time
import base64
from io import BytesIO
import pandas as pd
from PIL import Image
import streamlit as st

from utils.excel_handler import (
    load_student_data,
    suggest_default_column,
    validate_student_records,
    generate_sample_excel,
    generate_test_excel_with_errors
)
from utils.preview_renderer import (
    load_template_image,
    create_default_template_image,
    render_preview_overlay
)
from utils.font_manager import (
    AVAILABLE_FONT_FAMILIES,
    FONT_STYLES
)
from utils.generator import generate_batch_zip

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Bulk Certificate Generator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern styling and clean card interfaces
st.markdown("""
<style>
    /* Global Styles */
    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 24px 32px;
        border-radius: 14px;
        margin-bottom: 24px;
        border: 1px solid #334155;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
    }
    .header-title {
        color: #F8FAFC;
        font-size: 28px;
        font-weight: 700;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .header-subtitle {
        color: #94A3B8;
        font-size: 15px;
        margin-top: 6px;
    }
    .badge-pill {
        display: inline-block;
        background-color: #065F46;
        color: #34D399;
        font-size: 12px;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 20px;
        margin-top: 10px;
        border: 1px solid #059669;
    }
    .step-card {
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }
    .step-header {
        display: flex;
        align-items: center;
        gap: 10px;
        color: #F1F5F9;
        font-size: 18px;
        font-weight: 600;
        border-bottom: 1px solid #334155;
        padding-bottom: 8px;
        margin-bottom: 12px;
    }
    .step-num {
        background: #3B82F6;
        color: #FFFFFF;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: bold;
    }
    .metric-card {
        background: #0F172A;
        border: 1px solid #334155;
        border-radius: 8px;
        padding: 12px 16px;
        text-align: center;
    }
    .metric-title {
        color: #94A3B8;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-val {
        font-size: 24px;
        font-weight: 700;
        margin-top: 4px;
    }
    .val-good { color: #34D399; }
    .val-warn { color: #FBBF24; }
    .val-neutral { color: #60A5FA; }
    .meta-tag {
        display: inline-block;
        background-color: #334155;
        color: #E2E8F0;
        font-size: 13px;
        font-weight: 500;
        padding: 4px 10px;
        border-radius: 6px;
        margin-right: 6px;
        margin-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)


# Initialize Session State Variables
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "student_df" not in st.session_state:
    st.session_state.student_df = None
if "skipped_df" not in st.session_state:
    st.session_state.skipped_df = None
if "file_name" not in st.session_state:
    st.session_state.file_name = ""
if "validation_summary" not in st.session_state:
    st.session_state.validation_summary = {}

if "selected_name_col" not in st.session_state:
    st.session_state.selected_name_col = None
if "selected_roll_col" not in st.session_state:
    st.session_state.selected_roll_col = None
if "selected_extra_col" not in st.session_state:
    st.session_state.selected_extra_col = "(None)"

if "template_img" not in st.session_state:
    st.session_state.template_img = None
if "template_name" not in st.session_state:
    st.session_state.template_name = ""
if "template_meta" not in st.session_state:
    st.session_state.template_meta = None
if "generation_completed" not in st.session_state:
    st.session_state.generation_completed = False
if "combined_pdf_bytes" not in st.session_state:
    st.session_state.combined_pdf_bytes = None
if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None


# Helper class for wrapping file paths
class LocalFileWrapper:
    def __init__(self, filepath: str):
        self.name = os.path.basename(filepath)
        with open(filepath, "rb") as f:
            self._bytes = f.read()

    def getvalue(self):
        return self._bytes


def create_pdf_html_download_button(pdf_bytes: bytes, filename: str = "All_Student_Certificates.pdf", label: str = "📥 Direct PDF Download (.pdf)") -> str:
    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    return f"""
    <a href="data:application/pdf;base64,{b64}" download="{filename}" style="
        display: block;
        width: 100%;
        text-align: center;
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 13px;
        padding: 9px 14px;
        border-radius: 8px;
        text-decoration: none;
        border: 1px solid #3B82F6;
        box-shadow: 0 4px 10px rgba(29, 78, 216, 0.25);
        margin-top: 6px;
    ">
        {label}
    </a>
    """


# ==========================================
# SIDEBAR: Guidance, Privacy & Quick Tools
# ==========================================
with st.sidebar:
    st.markdown("### 🎓 Bulk Certificate Generator")
    st.caption("Version 1.2.0 • Local Edition")

    # Instant Downloads in sidebar if generated
    _sb_pdf_path = "All_Student_Certificates.pdf"
    _sb_zip_path = "certificates_bundle.zip"
    _has_pdf = os.path.exists(_sb_pdf_path)
    _has_zip = os.path.exists(_sb_zip_path)

    if _has_pdf or _has_zip:
        st.markdown("---")
        st.success("📄 **Certificates Ready!**")
        if _has_pdf:
            try:
                with open(_sb_pdf_path, "rb") as _f_pdf:
                    _sb_pdf_data = _f_pdf.read()
                st.download_button(
                    label="📄 Combined PDF (All in One)",
                    data=_sb_pdf_data,
                    file_name="All_Student_Certificates.pdf",
                    mime="application/pdf",
                    type="primary",
                    key="sidebar_download_combined_pdf_btn",
                    use_container_width=True
                )
                st.markdown(create_pdf_html_download_button(_sb_pdf_data, "All_Student_Certificates.pdf", "📥 Direct PDF Link (.pdf)"), unsafe_allow_html=True)
            except Exception:
                pass
        if _has_zip:
            try:
                with open(_sb_zip_path, "rb") as _f_zip:
                    _sb_zip_data = _f_zip.read()
                st.download_button(
                    label="📦 All Separate PDFs (.ZIP)",
                    data=_sb_zip_data,
                    file_name="certificates_bundle.zip",
                    mime="application/zip",
                    key="sidebar_download_zip_btn",
                    use_container_width=True
                )
            except Exception:
                pass

    st.markdown("---")
    st.markdown("#### 🔒 Privacy & Security")
    st.info(
        "**100% Local Processing**\n\n"
        "All Excel files, student records, and certificate templates are processed strictly inside your local computer's memory. "
        "No data is ever sent to external cloud servers or third parties."
    )

    st.markdown("---")
    st.markdown("#### 🚀 Quick Demos & Templates")

    # 1. Load CSAC Volunteering Template & 24-Student Roster
    csac_path = os.path.join("sample_data", "sreenidhi_certificate_template.png")
    if os.path.exists(csac_path):
        if st.button("🎖️ CSAC Guest Lecture Volunteering (33 Members)", use_container_width=True, type="primary"):
            sample_excel_path = os.path.join("sample_data", "csac_volunteers_clean.xlsx")
            if os.path.exists(sample_excel_path):
                st.session_state.raw_df = pd.read_excel(sample_excel_path, engine="openpyxl", dtype=object)
                st.session_state.file_name = "csac_volunteers_clean.xlsx"
                st.session_state.selected_name_col = "Student Name"
                st.session_state.selected_roll_col = "Roll Number"
                st.session_state.selected_extra_col = "Branch"
                st.session_state.generation_completed = False

            img, meta, _ = load_template_image(LocalFileWrapper(csac_path))
            st.session_state.template_img = img
            st.session_state.template_meta = meta
            st.session_state.template_name = "sreenidhi_certificate_template.png"
            
            # Preset optimal coordinates for this template (Name, Roll, Branch all in BOLD)
            st.session_state.name_x = 389
            st.session_state.name_y = 192
            st.session_state.name_size = 24
            st.session_state.name_color = "#1A2B4C"
            st.session_state.name_font = "Arial"
            st.session_state.name_style = "Bold"
            st.session_state.name_align = "center"
            
            st.session_state.roll_x = 365
            st.session_state.roll_y = 244
            st.session_state.roll_size = 13
            st.session_state.roll_color = "#1A2B4C"
            st.session_state.roll_font = "Arial"
            st.session_state.roll_style = "Bold"
            st.session_state.roll_prefix = ""
            st.session_state.roll_align = "center"

            st.session_state.selected_extra_col = "Branch"
            st.session_state.extra_x = 625
            st.session_state.extra_y = 244
            st.session_state.extra_size = 13
            st.session_state.extra_color = "#1A2B4C"
            st.session_state.extra_font = "Arial"
            st.session_state.extra_style = "Bold"
            st.session_state.extra_prefix = ""
            st.session_state.extra_align = "center"
            st.toast("✅ Loaded CSAC Guest Lecture (33 Members) Template!", icon="🎖️")
            st.rerun()

    # 2. Load Vector PDF Certificate Template (Landscape A4)
    pdf_sample_path = os.path.join("sample_data", "sample_certificate_template.pdf")
    if os.path.exists(pdf_sample_path):
        if st.button("📄 Vector PDF Template (Landscape A4)", use_container_width=True):
            img, meta, _ = load_template_image(LocalFileWrapper(pdf_sample_path))
            st.session_state.template_img = img
            st.session_state.template_meta = meta
            st.session_state.template_name = "sample_certificate_template.pdf"

            # Auto-center for high-res PDF
            w, h = img.size
            st.session_state.name_x = w // 2
            st.session_state.name_y = int(h * 0.43)
            st.session_state.name_size = 48
            st.session_state.name_color = "#0F172A"
            st.session_state.name_font = "Georgia"
            st.session_state.name_style = "Bold"

            st.session_state.roll_x = w // 2
            st.session_state.roll_y = int(h * 0.52)
            st.session_state.roll_size = 26
            st.session_state.roll_color = "#334155"
            st.session_state.roll_font = "Arial"
            st.session_state.roll_style = "Regular"
            st.session_state.roll_prefix = "Roll No: "
            st.toast("✅ Loaded Vector PDF Template!", icon="📄")
            st.rerun()

    # 3. Load Portrait Certificate Template
    portrait_sample_path = os.path.join("sample_data", "sample_portrait_template.png")
    if os.path.exists(portrait_sample_path):
        if st.button("📱 Academic Merit (Portrait PNG)", use_container_width=True):
            img, meta, _ = load_template_image(LocalFileWrapper(portrait_sample_path))
            st.session_state.template_img = img
            st.session_state.template_meta = meta
            st.session_state.template_name = "sample_portrait_template.png"

            # Auto-position for portrait
            w, h = img.size
            st.session_state.name_x = w // 2
            st.session_state.name_y = 390
            st.session_state.name_size = 36
            st.session_state.name_color = "#1E3A8A"
            st.session_state.name_font = "Georgia"
            st.session_state.name_style = "Bold"

            st.session_state.roll_x = w // 2
            st.session_state.roll_y = 445
            st.session_state.roll_size = 20
            st.session_state.roll_color = "#475569"
            st.session_state.roll_font = "Arial"
            st.session_state.roll_style = "Regular"
            st.session_state.roll_prefix = "Roll No: "
            st.toast("✅ Loaded Portrait Certificate Template!", icon="📱")
            st.rerun()

    # 4. Load Telugu Student Dataset
    telugu_excel_path = os.path.join("sample_data", "sample_students_telugu.xlsx")
    if os.path.exists(telugu_excel_path):
        if st.button("🇮🇳 Load Telugu Students (తెలుగు డేటా)", use_container_width=True):
            st.session_state.raw_df = pd.read_excel(telugu_excel_path, engine="openpyxl", dtype=object)
            st.session_state.file_name = "sample_students_telugu.xlsx"
            st.session_state.name_font = "Noto Sans Telugu"
            st.toast("✅ Loaded Telugu student dataset!", icon="🇮🇳")
            st.rerun()

    st.markdown("---")
    st.markdown("#### 📂 Download Sample Files")

    sample_bytes = generate_sample_excel()
    st.download_button(
        label="📄 Download Sample Excel Sheet",
        data=sample_bytes,
        file_name="sample_students.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if st.button("🔄 Reset / Clear All Data", use_container_width=True):
        st.session_state.raw_df = None
        st.session_state.student_df = None
        st.session_state.skipped_df = None
        st.session_state.file_name = ""
        st.session_state.validation_summary = {}
        st.session_state.template_img = None
        st.session_state.template_name = ""
        st.session_state.template_meta = None
        st.session_state.generation_completed = False
        st.session_state.combined_pdf_bytes = None
        st.session_state.zip_bytes = None
        st.rerun()

    st.markdown("---")
    st.markdown("#### 📋 Workflow Guide")
    st.markdown(
        """
        1. **Upload Spreadsheet**: `.xlsx` or `.csv`
        2. **Upload Template**: PDF, PNG, or JPG
        3. **Tune Coordinates**: Position fields & choose fonts
        4. **Preview Sample**: Verify layout with first record
        5. **Generate & Download**: Produce individual PDF certificates
        """
    )


# ==========================================
# MAIN HEADER
# ==========================================
st.markdown("""
<div class="main-header">
    <div class="header-title">
        <span>🎓</span> Bulk Certificate Generator
    </div>
    <div class="header-subtitle">
        Upload certificate templates (PDF, PNG, JPG), configure Telugu & English typography, and verify layouts without text collision.
    </div>
    <div>
        <span class="badge-pill">🛡️ 100% Offline & Private Processing</span>
        <span class="badge-pill" style="background-color: #1E3A8A; color: #60A5FA; border-color: #2563EB;">🎨 Multi-Format Vector & Raster Engine</span>
    </div>
</div>
""", unsafe_allow_html=True)

# Top Prominent Certificate Download Banners
_top_pdf_path = "All_Student_Certificates.pdf"
_top_zip_path = "certificates_bundle.zip"

_top_pdf_bytes = st.session_state.get("combined_pdf_bytes", None)
if _top_pdf_bytes is None and os.path.exists(_top_pdf_path):
    try:
        with open(_top_pdf_path, "rb") as _f:
            _top_pdf_bytes = _f.read()
            st.session_state.combined_pdf_bytes = _top_pdf_bytes
    except Exception:
        pass

_top_zip_bytes = st.session_state.get("zip_bytes", None)
if _top_zip_bytes is None and os.path.exists(_top_zip_path):
    try:
        with open(_top_zip_path, "rb") as _f:
            _top_zip_bytes = _f.read()
            st.session_state.zip_bytes = _top_zip_bytes
    except Exception:
        pass

if _top_pdf_bytes is not None or _top_zip_bytes is not None:
    _cert_count = 24
    if _top_zip_bytes is not None:
        try:
            import zipfile
            with zipfile.ZipFile(BytesIO(_top_zip_bytes)) as _zf:
                _cert_count = len([n for n in _zf.namelist() if n.endswith('.pdf')])
        except Exception:
            pass

    _pdf_size_mb = (len(_top_pdf_bytes) / (1024 * 1024)) if _top_pdf_bytes else 0.0
    _zip_size_mb = (len(_top_zip_bytes) / (1024 * 1024)) if _top_zip_bytes else 0.0

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #064E3B 0%, #065F46 100%); border: 2px solid #10B981; border-radius: 12px; padding: 18px 24px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.2);">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
            <div>
                <h3 style="color: #ECFDF5; margin: 0; font-size: 20px; font-weight: 700;">🎓 All {_cert_count} Certificates Generated & Ready!</h3>
                <p style="color: #A7F3D0; margin: 4px 0 0 0; font-size: 14px;">
                    Available as one sequential multi-page document (<b>All_Student_Certificates.pdf</b>, {_pdf_size_mb:.2f} MB, 1 student per page) and as an archive of individual PDFs (<b>certificates_bundle.zip</b>, {_zip_size_mb:.2f} MB).
                </p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    top_c1, top_c2 = st.columns([1, 1], gap="medium")
    with top_c1:
        if _top_pdf_bytes is not None:
            st.download_button(
                label="📄 DOWNLOAD COMBINED PDF (All_Student_Certificates.pdf)",
                data=_top_pdf_bytes,
                file_name="All_Student_Certificates.pdf",
                mime="application/pdf",
                type="primary",
                key="top_main_download_combined_pdf_btn",
                use_container_width=True
            )
            st.markdown(create_pdf_html_download_button(_top_pdf_bytes, "All_Student_Certificates.pdf", "📥 Direct Browser Download (.pdf extension)"), unsafe_allow_html=True)
    with top_c2:
        if _top_zip_bytes is not None:
            st.download_button(
                label="📦 DOWNLOAD INDIVIDUAL PDFS (.ZIP)",
                data=_top_zip_bytes,
                file_name="certificates_bundle.zip",
                mime="application/zip",
                key="top_main_download_zip_btn",
                use_container_width=True
            )
    st.markdown("<br>", unsafe_allow_html=True)


# ==========================================
# STEP 1: EXCEL / CSV SPREADSHEET UPLOAD & COLUMN MAPPING
# ==========================================
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <span class="step-num">1</span>
        <span>Student Data Import & Column Mapping</span>
    </div>
</div>
""", unsafe_allow_html=True)

upload_c1, upload_c2 = st.columns([1, 1], gap="medium")

with upload_c1:
    uploaded_file = st.file_uploader(
        "Upload Excel (.xlsx), CSV (.csv), or PDF student roster:",
        type=["xlsx", "csv", "xls", "pdf"],
        key="student_file_uploader",
        help="Upload files formatted with student names, roll numbers, and optional attributes."
    )

    if uploaded_file is not None:
        df_loaded, err_msg = load_student_data(uploaded_file)
        if err_msg:
            st.error(f"❌ {err_msg}")
        else:
            if st.session_state.file_name != uploaded_file.name:
                st.session_state.raw_df = df_loaded
                st.session_state.file_name = uploaded_file.name
                st.session_state.selected_name_col = None
                st.session_state.selected_roll_col = None
                st.session_state.selected_extra_col = "(None)"
                st.session_state.generation_completed = False
                st.toast(f"✅ Loaded {uploaded_file.name}", icon="📁")

    # Quick button to load the clean 31-student CSAC roster
    roster_path = os.path.join("sample_data", "csac_volunteers_clean.xlsx")
    if os.path.exists(roster_path):
        if st.button("🎖️ Use Clean CSAC Roster (31 Students - Verified)", use_container_width=True):
            st.session_state.raw_df = pd.read_excel(roster_path, engine="openpyxl", dtype=object)
            st.session_state.file_name = "csac_volunteers_clean.xlsx"
            st.session_state.selected_name_col = "Student Name"
            st.session_state.selected_roll_col = "Roll Number"
            st.session_state.selected_extra_col = "(None)"
            st.session_state.generation_completed = False
            st.rerun()

with upload_c2:
    if st.session_state.raw_df is not None:
        raw_df = st.session_state.raw_df
        st.markdown(f"**Loaded File:** `{st.session_state.file_name}`")
        st.markdown(f"**Raw Dataset:** `{len(raw_df)} total rows` × `{len(raw_df.columns)} columns`")
        st.caption("Raw data fields detected: " + ", ".join([f"`{c}`" for c in raw_df.columns]))
    else:
        st.info("ℹ️ No file uploaded yet. Upload a `.xlsx` or `.csv` file above, or select a demo from the sidebar.")


# Column selection and validation pipeline
if st.session_state.raw_df is not None:
    raw_df = st.session_state.raw_df
    columns_list = list(raw_df.columns)

    st.markdown("---")
    st.markdown("##### 🔀 Map Spreadsheet Columns")

    col_m1, col_m2, col_m3 = st.columns(3, gap="medium")

    # Smart column heuristics
    suggested_name = suggest_default_column(
        columns_list,
        ["name", "student_name", "candidate_name", "participant_name", "full_name", "student", "candidate", "participant", "learner"]
    ) or columns_list[0]

    suggested_roll = suggest_default_column(
        columns_list,
        ["roll", "roll_number", "roll_no", "id", "student_id", "reg_no", "registration_no", "registration_number", "enrollment", "urn", "usn", "uid"]
    ) or (columns_list[1] if len(columns_list) > 1 else columns_list[0])

    suggested_extra = suggest_default_column(
        columns_list,
        ["department", "course", "branch", "grade", "issue_date", "date", "event", "program"]
    ) or "(None)"

    curr_name_idx = columns_list.index(st.session_state.selected_name_col) if st.session_state.selected_name_col in columns_list else columns_list.index(suggested_name)
    curr_roll_idx = columns_list.index(st.session_state.selected_roll_col) if st.session_state.selected_roll_col in columns_list else columns_list.index(suggested_roll)
    
    extra_options = ["(None)"] + columns_list
    curr_extra_idx = extra_options.index(st.session_state.selected_extra_col) if st.session_state.selected_extra_col in extra_options else (extra_options.index(suggested_extra) if suggested_extra in extra_options else 0)

    with col_m1:
        selected_name = st.selectbox(
            "Student Name Column: *",
            options=columns_list,
            index=curr_name_idx,
            help="Select the column containing the student or participant's full name."
        )
        st.session_state.selected_name_col = selected_name

    with col_m2:
        selected_roll = st.selectbox(
            "Roll Number / ID Column: *",
            options=columns_list,
            index=curr_roll_idx,
            help="Select the column containing unique roll numbers or registration identifiers."
        )
        st.session_state.selected_roll_col = selected_roll

    with col_m3:
        selected_extra = st.selectbox(
            "Optional Extra Column:",
            options=extra_options,
            index=curr_extra_idx,
            help="Optional third field (e.g., Department, Course, Date, or Grade) to imprint on the certificate."
        )
        st.session_state.selected_extra_col = selected_extra

    # Run validation immediately with selected columns
    val_res = validate_student_records(
        raw_df,
        name_col=selected_name,
        roll_col=selected_roll,
        extra_col=selected_extra if selected_extra != "(None)" else None
    )

    valid_df = val_res["valid_df"]
    skipped_df = val_res["skipped_df"]

    # Store validated results in session_state for certificate generation & preview stages
    st.session_state.student_df = valid_df
    st.session_state.skipped_df = skipped_df
    st.session_state.validation_summary = val_res

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("##### 📊 Validation & Record Metrics")

    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Total Records in File</div>
            <div class="metric-val val-neutral">{val_res['total_rows']}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Valid Records (For Output)</div>
            <div class="metric-val val-good">{val_res['valid_count']}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col3:
        skipped_class = "val-warn" if val_res['skipped_count'] > 0 else "val-good"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Skipped / Invalid Rows</div>
            <div class="metric-val {skipped_class}">{val_res['skipped_count']}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col4:
        dup_class = "val-warn" if val_res['duplicate_count'] > 0 else "val-good"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Duplicate Roll Numbers</div>
            <div class="metric-val {dup_class}">{val_res['duplicate_count']}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Status Notification Box
    if val_res['skipped_count'] > 0:
        skip_breakdown = []
        if val_res['empty_name_count'] > 0:
            skip_breakdown.append(f"{val_res['empty_name_count']} missing Student Name")
        if val_res['empty_roll_count'] > 0:
            skip_breakdown.append(f"{val_res['empty_roll_count']} missing Roll Number")
        if val_res['duplicate_count'] > 0:
            skip_breakdown.append(f"{val_res['duplicate_count']} duplicate Roll Number")
        if val_res['empty_row_count'] > 0:
            skip_breakdown.append(f"{val_res['empty_row_count']} completely blank row")

        st.warning(
            f"⚠️ **{val_res['skipped_count']} rows were skipped to protect certificate integrity:** "
            + ", ".join(skip_breakdown)
            + ". Inspect the **'Skipped Rows'** tab below for detailed row-by-row reasons."
        )
    else:
        st.success(f"✅ **All {val_res['valid_count']} student records are valid!** No missing names, missing roll numbers, or duplicate entries detected.")

    tab_valid, tab_skipped, tab_raw = st.tabs([
        f"✅ Valid Records ({val_res['valid_count']})",
        f"⚠️ Skipped Rows ({val_res['skipped_count']})",
        f"📄 Raw Spreadsheet Data ({val_res['total_rows']})"
    ])

    with tab_valid:
        if val_res['valid_count'] > 0:
            display_valid = valid_df.copy()
            if "_excel_row" in display_valid.columns:
                cols = ["_excel_row"] + [c for c in display_valid.columns if c != "_excel_row"]
                display_valid = display_valid[cols].rename(columns={"_excel_row": "Excel Row #"})
            st.dataframe(display_valid, use_container_width=True, height=200)
            st.caption(f"Showing all {val_res['valid_count']} verified records. Student identity is strictly preserved.")
        else:
            st.error("No valid records found in the uploaded spreadsheet. Please check column mapping or upload a valid file.")

    with tab_skipped:
        if val_res['skipped_count'] > 0:
            st.dataframe(skipped_df, use_container_width=True, height=200)
            skipped_csv = skipped_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Skipped Rows Report (.CSV)",
                data=skipped_csv,
                file_name="skipped_rows_report.csv",
                mime="text/csv"
            )
        else:
            st.info("🎉 No rows were skipped! Every row in your spreadsheet passed all validation checks.")

    with tab_raw:
        st.dataframe(raw_df, use_container_width=True, height=200)

st.markdown("<br>", unsafe_allow_html=True)


# ==========================================
# STEP 2: CERTIFICATE TEMPLATE UPLOAD
# ==========================================
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <span class="step-num">2</span>
        <span>Certificate Template Upload (PDF, PNG, JPG)</span>
    </div>
</div>
""", unsafe_allow_html=True)

tmpl_col1, tmpl_col2 = st.columns([1, 1], gap="medium")

with tmpl_col1:
    uploaded_template = st.file_uploader(
        "Upload Certificate Template in PDF, PNG, or JPG format:",
        type=["pdf", "png", "jpg", "jpeg"],
        key="template_file_uploader",
        help="Upload high-resolution vector PDF, PNG, or JPG certificate templates. Landscape and Portrait orientations are fully supported."
    )

    if uploaded_template is not None:
        img, meta, err = load_template_image(uploaded_template)
        if err:
            st.error(f"❌ {err}")
        else:
            st.session_state.template_img = img
            st.session_state.template_meta = meta
            st.session_state.template_name = uploaded_template.name
            st.success(f"✅ Template loaded: {uploaded_template.name}")

    # Built-in template selector buttons
    st.markdown("###### Quick Template Presets:")
    pr_col1, pr_col2 = st.columns(2)

    with pr_col1:
        if os.path.exists(os.path.join("sample_data", "sreenidhi_certificate_template.png")):
            if st.button("🎖️ CSAC Volunteering (PNG)", use_container_width=True):
                img, meta, _ = load_template_image(LocalFileWrapper("sample_data/sreenidhi_certificate_template.png"))
                st.session_state.template_img = img
                st.session_state.template_meta = meta
                st.session_state.template_name = "sreenidhi_certificate_template.png"
                st.session_state.name_x = 435
                st.session_state.name_y = 224
                st.session_state.name_size = 20
                st.session_state.name_color = "#1A2B4C"
                st.session_state.roll_x = 345
                st.session_state.roll_y = 246
                st.session_state.roll_size = 11
                st.session_state.roll_color = "#1A2B4C"
                st.session_state.roll_prefix = ""
                st.rerun()

    with pr_col2:
        if os.path.exists(os.path.join("sample_data", "sample_certificate_template.pdf")):
            if st.button("📄 Vector Template (PDF)", use_container_width=True):
                img, meta, _ = load_template_image(LocalFileWrapper("sample_data/sample_certificate_template.pdf"))
                st.session_state.template_img = img
                st.session_state.template_meta = meta
                st.session_state.template_name = "sample_certificate_template.pdf"
                w, h = img.size
                st.session_state.name_x = w // 2
                st.session_state.name_y = int(h * 0.43)
                st.session_state.name_size = 48
                st.session_state.name_color = "#0F172A"
                st.session_state.roll_x = w // 2
                st.session_state.roll_y = int(h * 0.52)
                st.session_state.roll_size = 26
                st.session_state.roll_color = "#334155"
                st.session_state.roll_prefix = "Roll No: "
                st.rerun()

with tmpl_col2:
    if st.session_state.template_img is not None:
        t_img = st.session_state.template_img
        meta = st.session_state.template_meta or {}
        w, h = t_img.size
        orientation = meta.get("orientation", "Landscape" if w >= h else "Portrait")
        fmt = meta.get("format", "Image")
        aspect = meta.get("aspect_ratio", round(w / h, 2))

        # Specification Tags
        st.markdown(f"""
        <div>
            <span class="meta-tag">📄 Format: <b>{fmt}</b></span>
            <span class="meta-tag">📐 Orientation: <b>{orientation}</b></span>
            <span class="meta-tag">📏 Dimensions: <b>{w} × {h} px</b></span>
            <span class="meta-tag">⚖️ Ratio: <b>{aspect}:1</b></span>
        </div>
        """, unsafe_allow_html=True)

        st.image(t_img, caption=f"Template Preview: {st.session_state.template_name or 'Uploaded Template'}", use_container_width=True)
    else:
        st.info("ℹ️ No template uploaded yet. Please upload a PDF, PNG, or JPG certificate graphic, or choose a preset.")


st.markdown("<br>", unsafe_allow_html=True)


# ==========================================
# STEP 3 & 4: COORDINATE TUNING, TYPOGRAPHY & LIVE PREVIEW
# ==========================================
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <span class="step-num">3</span>
        <span>Layout Configuration, Typography & Real-Time Preview</span>
    </div>
</div>
""", unsafe_allow_html=True)

valid_df = st.session_state.student_df
t_img = st.session_state.template_img

if valid_df is None or len(valid_df) == 0 or t_img is None:
    st.warning("⚠️ Please provide valid student data in Step 1 and a certificate template in Step 2 to configure layout and live preview.")
else:
    img_w, img_h = t_img.size
    name_col = st.session_state.selected_name_col
    roll_col = st.session_state.selected_roll_col
    extra_col = st.session_state.selected_extra_col

    is_csac = "sreenidhi" in (st.session_state.template_name or "").lower()

    # Pre-populate defaults in session_state if not present
    if "name_x" not in st.session_state or st.session_state.name_x > img_w:
        st.session_state.name_x = min(389 if is_csac else img_w // 2, img_w)
    if "name_y" not in st.session_state or st.session_state.name_y > img_h:
        st.session_state.name_y = min(192 if is_csac else int(img_h * 0.42), img_h)
    if "name_size" not in st.session_state:
        st.session_state.name_size = 24 if is_csac else 48
    if "name_color" not in st.session_state:
        st.session_state.name_color = "#1A2B4C" if is_csac else "#0F172A"
    if "name_font" not in st.session_state:
        st.session_state.name_font = "Arial"
    if "name_style" not in st.session_state:
        st.session_state.name_style = "Bold"
    if "name_align" not in st.session_state:
        st.session_state.name_align = "center"

    if "roll_x" not in st.session_state or st.session_state.roll_x > img_w:
        st.session_state.roll_x = min(365 if is_csac else img_w // 2, img_w)
    if "roll_y" not in st.session_state or st.session_state.roll_y > img_h:
        st.session_state.roll_y = min(244 if is_csac else int(img_h * 0.52), img_h)
    if "roll_size" not in st.session_state:
        st.session_state.roll_size = 13 if is_csac else 24
    if "roll_color" not in st.session_state:
        st.session_state.roll_color = "#1A2B4C" if is_csac else "#334155"
    if "roll_font" not in st.session_state:
        st.session_state.roll_font = "Arial"
    if "roll_style" not in st.session_state:
        st.session_state.roll_style = "Bold" if is_csac else "Regular"
    if "roll_prefix" not in st.session_state:
        st.session_state.roll_prefix = "" if is_csac else "Roll No: "
    if "roll_align" not in st.session_state:
        st.session_state.roll_align = "center"

    # Callback helpers for precision coordinate controls
    def nudge_coord(key: str, delta: int, min_val: int, max_val: int):
        val = st.session_state.get(key, 0)
        st.session_state[key] = max(min_val, min(max_val, val + delta))

    def center_coord(key: str, target_val: int):
        st.session_state[key] = target_val

    cfg_col, prev_col = st.columns([1, 1], gap="large")

    with cfg_col:
        st.markdown("##### 🎨 Typography & Placement Controls")
        st.caption(f"Adjust position, Telugu/English fonts, and styles for `{name_col}` and `{roll_col}`:")

        tab_name, tab_roll, tab_extra = st.tabs(["👤 Student Name", "🆔 Roll Number", "🏷️ Extra Field"])

        with tab_name:
            col_x, col_y = st.columns(2)
            name_x = col_x.slider("X Position (px)", min_value=0, max_value=img_w, step=1, key="name_x")
            name_y = col_y.slider("Y Position (px)", min_value=0, max_value=img_h, step=1, key="name_y")

            # Quick Nudge Controls via on_click callbacks
            n1, n2, n3, n4, n5 = st.columns(5)
            n1.button("🎯 Center X", key="btn_center_name", on_click=center_coord, args=("name_x", img_w // 2), use_container_width=True)
            n2.button("⬆️ Up", key="btn_up_name", on_click=nudge_coord, args=("name_y", -5, 0, img_h), use_container_width=True)
            n3.button("⬇️ Down", key="btn_down_name", on_click=nudge_coord, args=("name_y", 5, 0, img_h), use_container_width=True)
            n4.button("⬅️ Left", key="btn_left_name", on_click=nudge_coord, args=("name_x", -5, 0, img_w), use_container_width=True)
            n5.button("➡️ Right", key="btn_right_name", on_click=nudge_coord, args=("name_x", 5, 0, img_w), use_container_width=True)

            col_f, col_st = st.columns(2)
            name_font = col_f.selectbox("Font Family", AVAILABLE_FONT_FAMILIES, key="name_font", help="Supports English and Telugu Unicode (Noto Sans Telugu / Nirmala UI)")
            name_style = col_st.selectbox("Font Style", FONT_STYLES, key="name_style")

            col_sz, col_cl, col_al = st.columns(3)
            name_size = col_sz.slider("Font Size", min_value=8, max_value=140, step=1, key="name_size")
            name_color = col_cl.color_picker("Text Color", key="name_color")
            name_align = col_al.selectbox("Alignment", ["center", "left", "right"], key="name_align")

        with tab_roll:
            col_rx, col_ry = st.columns(2)
            roll_x = col_rx.slider("X Position (px)", min_value=0, max_value=img_w, step=1, key="roll_x")
            roll_y = col_ry.slider("Y Position (px)", min_value=0, max_value=img_h, step=1, key="roll_y")

            # Quick Nudge Controls via on_click callbacks
            rn1, rn2, rn3, rn4, rn5 = st.columns(5)
            rn1.button("🎯 Center X", key="btn_center_roll", on_click=center_coord, args=("roll_x", img_w // 2), use_container_width=True)
            rn2.button("⬆️ Up", key="btn_up_roll", on_click=nudge_coord, args=("roll_y", -5, 0, img_h), use_container_width=True)
            rn3.button("⬇️ Down", key="btn_down_roll", on_click=nudge_coord, args=("roll_y", 5, 0, img_h), use_container_width=True)
            rn4.button("⬅️ Left", key="btn_left_roll", on_click=nudge_coord, args=("roll_x", -5, 0, img_w), use_container_width=True)
            rn5.button("➡️ Right", key="btn_right_roll", on_click=nudge_coord, args=("roll_x", 5, 0, img_w), use_container_width=True)

            col_rf, col_rst = st.columns(2)
            roll_font = col_rf.selectbox("Font Family", AVAILABLE_FONT_FAMILIES, key="roll_font")
            roll_style = col_rst.selectbox("Font Style", FONT_STYLES, key="roll_style")

            col_rsz, col_rcl, col_ral = st.columns(3)
            roll_size = col_rsz.slider("Font Size", min_value=8, max_value=100, step=1, key="roll_size")
            roll_color = col_rcl.color_picker("Text Color", key="roll_color")
            roll_align = col_ral.selectbox("Alignment", ["center", "left", "right"], key="roll_align")

            roll_prefix = st.text_input("Prefix Text (Optional)", key="roll_prefix")

        with tab_extra:
            if extra_col == "(None)":
                st.caption("No extra column selected in Step 1.")
                extra_x, extra_y, extra_size, extra_color, extra_prefix, extra_align, extra_font, extra_style = (
                    img_w // 2, int(img_h * 0.58), 22, "#475569", "", "center", "Arial", "Regular"
                )
            else:
                col_ex, col_ey = st.columns(2)
                extra_x = col_ex.slider("X Position (px)", min_value=0, max_value=img_w, value=img_w // 2, step=1, key="extra_x")
                extra_y = col_ey.slider("Y Position (px)", min_value=0, max_value=img_h, value=int(img_h * 0.58), step=1, key="extra_y")

                col_ef, col_est = st.columns(2)
                extra_font = col_ef.selectbox("Font Family", AVAILABLE_FONT_FAMILIES, key="extra_font")
                extra_style = col_est.selectbox("Font Style", FONT_STYLES, key="extra_style")

                col_esz, col_ecl, col_eal = st.columns(3)
                extra_size = col_esz.slider("Font Size", min_value=8, max_value=100, value=22, step=1, key="extra_size")
                extra_color = col_ecl.color_picker("Text Color", value="#475569", key="extra_color")
                extra_align = col_eal.selectbox("Alignment", ["center", "left", "right"], index=0, key="extra_align")
                extra_prefix = st.text_input("Prefix Text", value=f"{extra_col}: ", key="extra_prefix")

    # Live Preview Area
    with prev_col:
        st.markdown("##### 👁️ Live Certificate Preview")

        # Select record to preview: default is the first valid record (#1)
        prev_c1, prev_c2 = st.columns([2, 1])
        preview_idx = prev_c1.slider("Record to Preview (Default: Record #1)", min_value=1, max_value=len(valid_df), value=1, step=1)
        selected_record = valid_df.iloc[preview_idx - 1]

        show_guides = prev_c2.checkbox("Show Guides", value=True, help="Toggle bounding boxes and coordinate crosshairs")

        # Construct preview fields with strict raw preservation
        sample_name = str(selected_record.get(name_col, "Student Name"))
        raw_roll_val = selected_record.get(roll_col, "CS-001")
        sample_roll = f"{roll_prefix}{raw_roll_val}"
        
        fields_config = {
            "name": {
                "text": sample_name,
                "x": name_x,
                "y": name_y,
                "font_size": name_size,
                "font_family": name_font,
                "font_style": name_style,
                "color": name_color,
                "align": name_align,
                "label": "Student Name"
            },
            "roll": {
                "text": sample_roll,
                "x": roll_x,
                "y": roll_y,
                "font_size": roll_size,
                "font_family": roll_font,
                "font_style": roll_style,
                "color": roll_color,
                "align": roll_align,
                "label": "Roll Number"
            }
        }

        if extra_col != "(None)" and extra_col in selected_record:
            raw_extra_val = selected_record.get(extra_col, "")
            fields_config["extra"] = {
                "text": f"{extra_prefix}{raw_extra_val}",
                "x": extra_x,
                "y": extra_y,
                "font_size": extra_size,
                "font_family": extra_font,
                "font_style": extra_style,
                "color": extra_color,
                "align": extra_align,
                "label": extra_col
            }

        # Render preview with intelligent text overlap and boundary collision detection
        preview_image, layout_warnings, bboxes = render_preview_overlay(
            t_img,
            fields_config,
            show_bounding_box=show_guides,
            show_anchors=show_guides
        )

        # Overlap / collision alert banner
        if layout_warnings:
            for w in layout_warnings:
                st.warning(w)
        else:
            st.success("✅ **Layout Clear:** No text collisions or boundary overflows detected.")

        orig_row = selected_record.get('_excel_row', preview_idx + 1)
        st.image(
            preview_image,
            caption=f"Previewing Valid Record #{preview_idx} of {len(valid_df)} (Excel Row #{orig_row}): {sample_name}",
            use_container_width=True
        )

        # Coordinates & bounding box summary expander
        with st.expander("📐 Technical Coordinate & Bounding Box Details"):
            for fk, bb in bboxes.items():
                lbl = fields_config[fk]["label"]
                bw, bh = bb[2] - bb[0], bb[3] - bb[1]
                st.markdown(f"- **{lbl}:** `Box: {bw} × {bh} px` | `X: [{bb[0]}, {bb[2]}]` | `Y: [{bb[1]}, {bb[3]}]`")


st.markdown("<br>", unsafe_allow_html=True)


# ==========================================
# STEP 5: GENERATION & VERIFICATION PIPELINE
# ==========================================
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <span class="step-num">4</span>
        <span>Certificate Batch Generation</span>
    </div>
</div>
""", unsafe_allow_html=True)

gen_col1, gen_col2 = st.columns([1, 1], gap="medium")

with gen_col1:
    st.markdown("##### 📋 Pre-Flight Checklist")

    check_data = st.session_state.student_df is not None and len(st.session_state.student_df) > 0
    check_template = st.session_state.template_img is not None
    check_ready = check_data and check_template

    total_valid_count = len(st.session_state.student_df) if check_data else 0

    st.markdown(f"{'✅' if check_data else '❌'} **Verified Student Records:** {f'{total_valid_count} ready' if check_data else 'No valid records'}")
    st.markdown(f"{'✅' if check_template else '❌'} **Certificate Template:** {st.session_state.template_name or 'Ready' if check_template else 'Missing'}")
    st.markdown(f"{'✅' if check_ready else '❌'} **Output Engine:** Individual High-Resolution PDFs packaged in a ZIP")

    if not check_ready:
        st.warning("⚠️ Please provide validated spreadsheet records and a certificate template above.")
    else:
        st.success(f"🎯 **{total_valid_count} valid records** confirmed for batch rendering!")

with gen_col2:
    st.markdown("##### ⚙️ Generation Pipeline")
    st.caption("All certificate generation takes place locally on your computer with zero cloud exposure.")

    filename_pattern = st.selectbox(
        "PDF Filename Convention:",
        ["RollNumber_StudentName.pdf", "StudentName_RollNumber.pdf", "StudentName.pdf", "Certificate_RollNumber.pdf"],
        index=0,
        help="Naming pattern for each generated PDF inside the ZIP archive."
    )

    if st.button("🚀 Generate All Certificates (PDF ZIP Archive)", disabled=not check_ready, use_container_width=True, type="primary"):
        progress_bar = st.progress(0, text="Starting batch PDF certificate generation...")

        def on_gen_progress(current_i, total_n, student_name):
            pct = int((current_i / total_n) * 100)
            progress_bar.progress(pct, text=f"Rendering certificate {current_i}/{total_n}: {student_name}...")

        # Construct generation fields config
        gen_fields_config = {
            "name": {
                "x": st.session_state.name_x,
                "y": st.session_state.name_y,
                "font_size": st.session_state.name_size,
                "font_family": st.session_state.name_font,
                "font_style": st.session_state.name_style,
                "color": st.session_state.name_color,
                "align": st.session_state.name_align
            },
            "roll": {
                "x": st.session_state.roll_x,
                "y": st.session_state.roll_y,
                "font_size": st.session_state.roll_size,
                "font_family": st.session_state.roll_font,
                "font_style": st.session_state.roll_style,
                "color": st.session_state.roll_color,
                "align": st.session_state.roll_align,
                "prefix": st.session_state.roll_prefix
            }
        }

        if st.session_state.selected_extra_col != "(None)" and "extra_x" in st.session_state:
            gen_fields_config["extra"] = {
                "x": st.session_state.extra_x,
                "y": st.session_state.extra_y,
                "font_size": st.session_state.extra_size,
                "font_family": st.session_state.get("extra_font", "Arial"),
                "font_style": st.session_state.get("extra_style", "Regular"),
                "color": st.session_state.get("extra_color", "#475569"),
                "align": st.session_state.get("extra_align", "center"),
                "prefix": st.session_state.get("extra_prefix", "")
            }

        zip_bytes, combined_pdf_bytes, manifest = generate_batch_zip(
            base_template=st.session_state.template_img,
            valid_df=st.session_state.student_df,
            name_col=st.session_state.selected_name_col,
            roll_col=st.session_state.selected_roll_col,
            extra_col=st.session_state.selected_extra_col if st.session_state.selected_extra_col != "(None)" else None,
            fields_config=gen_fields_config,
            naming_pattern=filename_pattern,
            progress_callback=on_gen_progress
        )

        st.session_state.zip_bytes = zip_bytes
        st.session_state.combined_pdf_bytes = combined_pdf_bytes
        st.session_state.zip_manifest = manifest
        st.session_state.generation_completed = True

        # Save to disk as well
        with open("certificates_bundle.zip", "wb") as f:
            f.write(zip_bytes)

        with open("All_Student_Certificates.pdf", "wb") as f:
            f.write(combined_pdf_bytes)

        st.balloons()
        st.success(f"🎉 Successfully rendered {len(manifest)} certificates! Download the combined multi-page PDF or individual ZIP archive below.")


st.markdown("<br>", unsafe_allow_html=True)


# ==========================================
# STEP 6: DOWNLOAD ARCHIVE
# ==========================================
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <span class="step-num">5</span>
        <span>Export & Download</span>
    </div>
</div>
""", unsafe_allow_html=True)

# Check if combined_pdf_bytes exists in session state or on disk
pdf_to_download = st.session_state.get("combined_pdf_bytes", None)
if pdf_to_download is None and os.path.exists("All_Student_Certificates.pdf"):
    with open("All_Student_Certificates.pdf", "rb") as f:
        pdf_to_download = f.read()

# Check if zip_bytes exists in session state or on disk
zip_to_download = st.session_state.get("zip_bytes", None)
if zip_to_download is None and os.path.exists("certificates_bundle.zip"):
    with open("certificates_bundle.zip", "rb") as f:
        zip_to_download = f.read()

total_valid = len(st.session_state.student_df) if st.session_state.student_df is not None else 24

down_col1, down_col2 = st.columns([1, 1], gap="medium")

with down_col1:
    st.markdown("##### 📄 Option 1: Combined Multi-Page PDF")
    if pdf_to_download is not None:
        pdf_size_kb = len(pdf_to_download) / 1024
        st.success(f"✅ Combined document ready: `{total_valid}` pages ({pdf_size_kb:.1f} KB).")
        st.caption("1 student certificate per page in the exact spreadsheet order.")
        st.download_button(
            label=f"📄 Download All_Student_Certificates.pdf ({total_valid} Pages)",
            data=pdf_to_download,
            file_name="All_Student_Certificates.pdf",
            mime="application/pdf",
            type="primary",
            key="step5_download_combined_pdf_btn",
            use_container_width=True
        )
        st.markdown(create_pdf_html_download_button(pdf_to_download, "All_Student_Certificates.pdf", f"📥 Direct Browser Download ({total_valid} Pages .pdf)"), unsafe_allow_html=True)
    else:
        st.info("ℹ️ Generate certificates in Step 4 to produce your combined PDF.")

with down_col2:
    st.markdown("##### 📦 Option 2: Individual Certificates (.ZIP)")
    if zip_to_download is not None:
        zip_size_kb = len(zip_to_download) / 1024
        st.info(f"✅ ZIP archive ready: `{total_valid}` separate PDFs ({zip_size_kb:.1f} KB).")
        st.caption("Individual PDF files named by student roll number and name.")
        st.download_button(
            label=f"📥 Download Certificates Bundle ({total_valid} PDFs in .ZIP)",
            data=zip_to_download,
            file_name="certificates_bundle.zip",
            mime="application/zip",
            key="step5_download_zip_btn",
            use_container_width=True
        )
    else:
        st.info("ℹ️ Generate certificates in Step 4 to produce your ZIP download.")

st.markdown("---")
st.markdown("##### 📌 Batch Summary & Verification")
if st.session_state.student_df is not None:
    val_count = len(st.session_state.student_df)
    skip_count = len(st.session_state.skipped_df) if st.session_state.skipped_df is not None else 0
    sm_c1, sm_c2, sm_c3 = st.columns(3)
    with sm_c1:
        st.markdown(f"- **Total Records:** `{val_count}`")
        st.markdown(f"- **Skipped Records:** `{skip_count}`")
    with sm_c2:
        st.markdown(f"- **Combined PDF:** `All_Student_Certificates.pdf`")
        st.markdown(f"- **Pages in Combined:** `{val_count} pages (1 per student)`")
    with sm_c3:
        st.markdown(f"- **Archive File:** `certificates_bundle.zip`")
        st.markdown(f"- **Privacy Status:** `100% Local (RAM Only)`")
else:
    st.caption("Upload student data to view batch metrics.")
