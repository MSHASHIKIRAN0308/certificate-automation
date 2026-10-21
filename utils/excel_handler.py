"""
Excel and CSV data processing and validation utility for Bulk Certificate Generator.
"""
from io import BytesIO
from typing import Tuple, List, Optional, Dict, Any
import pandas as pd


def load_student_data(uploaded_file) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Safely load an uploaded Excel (.xlsx) or CSV (.csv) file into a pandas DataFrame.
    Preserves raw data without silent modifications to strings, leading zeros, or spacing.
    Returns: (DataFrame, error_message)
    """
    if uploaded_file is None:
        return None, "No file uploaded."

    try:
        filename = uploaded_file.name.lower()
        file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

        if filename.endswith(".csv"):
            # Try reading with utf-8 first, fallback to latin-1 / cp1252 for broad CSV compatibility
            try:
                df = pd.read_csv(BytesIO(file_bytes), dtype=object, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(BytesIO(file_bytes), dtype=object, encoding="latin-1")
        elif filename.endswith((".xlsx", ".xls")):
            # Explicitly use openpyxl for .xlsx and read as object to prevent float/int distortion
            df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl", dtype=object)
        elif filename.endswith(".pdf"):
            # Extract structured student roster from PDF
            import pymupdf, re
            doc = pymupdf.open(stream=file_bytes, filetype="pdf")
            records = []
            for page in doc:
                words = page.get_text("words")
                line_groups = {}
                for w in words:
                    y_approx = round(w[1], -1)
                    line_groups.setdefault(y_approx, []).append(w)
                for y in sorted(line_groups.keys()):
                    items = sorted(line_groups[y], key=lambda x: x[0])
                    line_text = " ".join([it[4] for it in items])
                    if any(k in line_text.lower() for k in ["s.no", "roll no", "roll_no"]):
                        continue
                    m = re.match(r"^(\d+)\s+([0-9a-zA-Z]+)\s+(.+)$", line_text)
                    if m:
                        s_no = m.group(1)
                        roll = m.group(2)
                        rest = m.group(3).strip()
                        m_days = re.search(r"\s+(\d+)$", rest)
                        days = m_days.group(1) if m_days else ""
                        name_part = rest[:m_days.start()].strip() if m_days else rest
                        m_branch = re.search(r"\s+(I{1,3}\s*Year|IV\s*Year|IV|III|II)\s+([A-Za-z0-9\-]+)(?:\s+([A-Za-z0-9]+))?$", name_part)
                        if m_branch:
                            name = name_part[:m_branch.start()].strip()
                            year = m_branch.group(1).strip()
                            branch = m_branch.group(2).strip()
                            sec = m_branch.group(3).strip() if m_branch.group(3) else ""
                        else:
                            name = name_part
                            year, branch, sec = "", "", ""
                        records.append({
                            "S.No": s_no,
                            "Roll No": roll,
                            "Name": name,
                            "Year": year,
                            "Branch": branch,
                            "Section": sec,
                            "No of days": days
                        })
            doc.close()
            df = pd.DataFrame(records)
        else:
            return None, "Unsupported file format. Please upload a .xlsx, .csv, or .pdf roster file."

        if df.empty:
            return None, "The uploaded file contains no data rows."

        # Keep column names clean for lookup, but strip only header whitespace
        df.columns = [str(col).strip() for col in df.columns]

        if len(df.columns) == 0:
            return None, "No columns detected in the uploaded file."

        return df, None
    except Exception as e:
        return None, f"Failed to read file: {str(e)}"


def suggest_default_column(columns: List[str], keywords: List[str]) -> Optional[str]:
    """
    Heuristically matches common column names based on a prioritized list of keywords.
    Handles varied column naming conventions across different institutions.
    """
    col_clean_map = {
        col.lower().replace("_", "").replace(" ", "").replace("-", "").replace(".", ""): col
        for col in columns
    }

    # 1. Exact match check
    for kw in keywords:
        clean_kw = kw.lower().replace("_", "").replace(" ", "").replace("-", "").replace(".", "")
        if clean_kw in col_clean_map:
            return col_clean_map[clean_kw]

    # 2. Substring match check
    for kw in keywords:
        clean_kw = kw.lower().replace("_", "").replace(" ", "").replace("-", "").replace(".", "")
        for key, orig_col in col_clean_map.items():
            if clean_kw in key:
                return orig_col

    return None


def is_empty_value(val: Any) -> bool:
    """
    Checks if a field value is considered empty, null, or whitespace-only.
    Does NOT modify the value.
    """
    if val is None or pd.isna(val):
        return True
    str_val = str(val).strip()
    if str_val == "" or str_val.lower() in ("nan", "none", "null", "n/a", "na", "<na>"):
        return True
    return False


def validate_student_records(
    df: pd.DataFrame,
    name_col: str,
    roll_col: str,
    extra_col: Optional[str] = None
) -> Dict[str, Any]:
    """
    Validates student records against:
    - Missing/empty student names
    - Missing/empty roll numbers
    - Duplicate roll numbers (or identical duplicate records)
    - Entirely empty rows

    IMPORTANT: Does NOT silently modify student names or roll numbers.
    Original raw values are preserved in valid_df exactly as entered by the user.

    Returns a dictionary containing:
    - 'valid_df': DataFrame with valid student records
    - 'skipped_df': DataFrame with skipped records, including original row number and reason
    - 'total_rows': Total rows in the uploaded spreadsheet
    - 'valid_count': Count of valid rows
    - 'skipped_count': Count of skipped rows
    - 'empty_name_count': Count of rows with missing names
    - 'empty_roll_count': Count of rows with missing roll numbers
    - 'duplicate_count': Count of duplicate records
    - 'empty_row_count': Count of completely blank rows
    """
    if df is None or df.empty:
        return {
            "valid_df": pd.DataFrame(),
            "skipped_df": pd.DataFrame(),
            "total_rows": 0,
            "valid_count": 0,
            "skipped_count": 0,
            "empty_name_count": 0,
            "empty_roll_count": 0,
            "duplicate_count": 0,
            "empty_row_count": 0,
        }

    valid_rows = []
    skipped_rows = []

    seen_roll_numbers = {}  # roll_str -> first_seen_excel_row
    
    empty_name_count = 0
    empty_roll_count = 0
    duplicate_count = 0
    empty_row_count = 0

    for idx, row in df.iterrows():
        excel_row_num = idx + 2  # 1-indexed Excel row (Header is row 1)
        
        # Check for completely blank row across all columns
        all_empty = all(is_empty_value(row[c]) for c in df.columns)
        if all_empty:
            empty_row_count += 1
            skipped_rows.append({
                "Excel Row": excel_row_num,
                "Student Name": "[Empty]",
                "Roll Number": "[Empty]",
                "Skip Reason": "Completely blank row"
            })
            continue

        raw_name = row.get(name_col, None)
        raw_roll = row.get(roll_col, None)

        name_empty = is_empty_value(raw_name)
        roll_empty = is_empty_value(raw_roll)

        if name_empty and roll_empty:
            empty_name_count += 1
            empty_roll_count += 1
            skipped_rows.append({
                "Excel Row": excel_row_num,
                "Student Name": "[Missing]",
                "Roll Number": "[Missing]",
                "Skip Reason": "Both Student Name and Roll Number are missing"
            })
            continue

        if name_empty:
            empty_name_count += 1
            skipped_rows.append({
                "Excel Row": excel_row_num,
                "Student Name": "[Missing]",
                "Roll Number": str(raw_roll) if raw_roll is not None else "",
                "Skip Reason": f"Missing or empty Student Name in column '{name_col}'"
            })
            continue

        if roll_empty:
            empty_roll_count += 1
            skipped_rows.append({
                "Excel Row": excel_row_num,
                "Student Name": str(raw_name) if raw_name is not None else "",
                "Roll Number": "[Missing]",
                "Skip Reason": f"Missing or empty Roll Number in column '{roll_col}'"
            })
            continue

        # Check for duplicate roll numbers
        # Use trimmed string for duplicate detection key only, do NOT alter original raw_roll or raw_name
        roll_key = str(raw_roll).strip().lower()
        if roll_key in seen_roll_numbers:
            duplicate_count += 1
            first_seen_row = seen_roll_numbers[roll_key]
            skipped_rows.append({
                "Excel Row": excel_row_num,
                "Student Name": str(raw_name),
                "Roll Number": str(raw_roll),
                "Skip Reason": f"Duplicate Roll Number '{raw_roll}' (First seen at Excel row #{first_seen_row})"
            })
            continue

        # Record is valid!
        seen_roll_numbers[roll_key] = excel_row_num
        valid_record = dict(row)
        valid_record["_excel_row"] = excel_row_num
        valid_rows.append(valid_record)

    valid_df = pd.DataFrame(valid_rows) if valid_rows else pd.DataFrame(columns=list(df.columns) + ["_excel_row"])
    skipped_df = pd.DataFrame(skipped_rows) if skipped_rows else pd.DataFrame(columns=["Excel Row", "Student Name", "Roll Number", "Skip Reason"])

    return {
        "valid_df": valid_df,
        "skipped_df": skipped_df,
        "total_rows": len(df),
        "valid_count": len(valid_df),
        "skipped_count": len(skipped_df),
        "empty_name_count": empty_name_count,
        "empty_roll_count": empty_roll_count,
        "duplicate_count": duplicate_count,
        "empty_row_count": empty_row_count,
    }


def generate_sample_excel() -> bytes:
    """
    Generates an in-memory sample Excel file (.xlsx) with diverse student records for testing.
    Includes edge test cases: international characters, standard formatting, and extra columns.
    """
    sample_records = [
        {"Student Name": "Sophia Montgomery", "Roll Number": "CS2026-001", "Department": "Computer Science", "Issue Date": "16-Oct-2026", "Grade": "Distinction"},
        {"Student Name": "Alexander Chen", "Roll Number": "CS2026-002", "Department": "Information Technology", "Issue Date": "16-Oct-2026", "Grade": "Distinction"},
        {"Student Name": "Aarav Sharma", "Roll Number": "EE2026-015", "Department": "Electrical Engineering", "Issue Date": "16-Oct-2026", "Grade": "First Class"},
        {"Student Name": "Isabella Rossi", "Roll Number": "ME2026-042", "Department": "Mechanical Engineering", "Issue Date": "16-Oct-2026", "Grade": "Distinction"},
        {"Student Name": "Liam O'Connor", "Roll Number": "DS2026-019", "Department": "Data Science & AI", "Issue Date": "16-Oct-2026", "Grade": "First Class"},
        {"Student Name": "Renée Dupré", "Roll Number": "BT2026-033", "Department": "Biotechnology", "Issue Date": "16-Oct-2026", "Grade": "Distinction"},
    ]
    df = pd.DataFrame(sample_records)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Students")
    return output.getvalue()


def generate_test_excel_with_errors() -> bytes:
    """
    Generates an Excel spreadsheet containing intentional edge cases for validation testing:
    - Valid records
    - Missing name
    - Missing roll number
    - Duplicate roll number
    - Completely blank row
    - Name with whitespace / apostrophe to verify preservation
    """
    test_records = [
        {"Student Name": "Marcus Vance", "Roll Number": "REG-1001", "Department": "Design"},
        {"Student Name": "Elena Rostova", "Roll Number": "REG-1002", "Department": "Robotics"},
        {"Student Name": "", "Roll Number": "REG-1003", "Department": "Design"},  # Empty Name
        {"Student Name": "David Kim", "Roll Number": "", "Department": "Robotics"},  # Empty Roll
        {"Student Name": "Marcus Vance Duplicate", "Roll Number": "REG-1001", "Department": "Design"},  # Duplicate Roll
        {"Student Name": None, "Roll Number": None, "Department": None},  # Blank row
        {"Student Name": "  O'Neil MacDonald  ", "Roll Number": "00987", "Department": "Architecture"},  # Preserved spaces & leading zero
    ]
    df = pd.DataFrame(test_records)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Students")
    return output.getvalue()
