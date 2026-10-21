# 🎓 Bulk Certificate Generator

A complete, modern, and privacy-first local web application built with **Python** and **Streamlit** to generate individual personalized PDF certificates from Excel spreadsheets.

---

## 🔒 100% Local & Privacy-Guaranteed
All student data, names, roll numbers, and certificate templates are processed strictly inside your local computer's memory. No data is ever transmitted to external servers or cloud APIs.

---

## 🚀 Key Features

1. **Spreadsheet Upload & Parsing**:
   - Supports `.xlsx` (via `openpyxl`), `.xls`, and `.csv` (with UTF-8 and Latin-1 fallback).
   - Zero silent mutations: Leading zeros, spacing, and accented characters are strictly preserved.
   - Comprehensive validation: Detects empty names, empty roll numbers, duplicate roll numbers, and blank rows.
   - Real-time record metrics: Total Rows, Valid Records, Skipped Rows, and Duplicate Count.
   - Row-by-row diagnostic report for skipped entries with downloadable CSV report.

2. **Multi-Format Certificate Template Engine**:
   - Supports vector `.pdf`, high-res `.png`, and `.jpg` / `.jpeg` templates.
   - Non-destructive processing: Preserves original designs, vector lines, logos, institutional crests, and signatures without flattening.
   - Handles **Landscape** and **Portrait** orientations with automatic dimension detection and aspect ratio calculation.

3. **Multi-Language Typography & Telugu Font Support**:
   - Bundled **Noto Sans Telugu** (Regular & Bold) and Windows **Nirmala UI** for flawless Telugu Unicode script rendering (e.g., `రాహుల్ వర్మ`).
   - English typography: Arial, Calibri, Times New Roman, and Georgia.
   - Full styling options: Font Family, Font Style (Regular, Bold, Italic, Bold Italic), Font Size, Hex Color Picker, and Alignment (Left, Center, Right).

4. **Coordinate Placement & Overlap Collision Detection**:
   - Pixel-accurate X and Y position sliders with quick-nudge buttons (`🎯 Center X`, `⬆️ Up`, `⬇️ Down`, `⬅️ Left`, `➡️ Right`).
   - Dynamic boundary calculation: Prevents text from spilling outside certificate edges.
   - Automated collision engine: Warns immediately if Student Name and Roll Number text bounding boxes collide or overlap.
   - Sample preview bound to the first valid spreadsheet record with toggleable crosshairs and bounding boxes.

5. **Pre-flight Validation & Batch Generation**:
   - Automated checklist verifying validated record count and template readiness.
   - Simulation engine with live progress bars and status feedback.

6. **Batch ZIP Export**:
   - Single-click packaging of individual PDF certificates into a downloadable `.zip` archive.

---

## 📁 Project Structure

```text
certficate automation/
│
├── app.py                     # Main Streamlit application
├── requirements.txt           # Python dependencies
├── create_samples.py          # Script to generate sample test assets
├── test_excel_validation.py   # Test suite for Excel/CSV upload & validation
├── README.md                  # Documentation and setup guide
│
├── utils/
│   ├── __init__.py
│   ├── excel_handler.py       # Excel/CSV parser, validator & column mapping
│   └── preview_renderer.py    # Template loader and live preview overlay
│
└── sample_data/
    ├── sample_students.xlsx             # Clean sample student records (6 records)
    ├── sample_students.csv              # CSV format sample student records
    ├── sample_students_with_errors.xlsx # Edge-case dataset with intentional errors
    └── sample_template.png              # Modern sample certificate background
```

---

## 🛠️ Installation & Setup

### 1. Prerequisites
Ensure you have **Python 3.10+** installed on your system.

### 2. Install Dependencies
Open your terminal or PowerShell in the project directory:

```bash
pip install -r requirements.txt
```

---

## ▶️ Running the Application

Launch the Streamlit web application with:

```bash
streamlit run app.py
```

The application will automatically open in your default web browser at:
`http://localhost:8501`

---

## 💡 Quick Start Demo
If you don't have an Excel sheet or template handy:
1. Launch the app.
2. In the left sidebar, click **"📥 Load Sample Excel & Template"**.
3. Explore the live preview, adjust the coordinate sliders, and test the validation run!
