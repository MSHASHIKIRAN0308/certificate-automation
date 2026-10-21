"""
Comprehensive unit and integration tests for Excel/CSV upload, validation, and column mapping.
"""
from io import BytesIO
import pandas as pd

from utils.excel_handler import (
    load_student_data,
    suggest_default_column,
    validate_student_records,
    generate_test_excel_with_errors
)


class DummyUploadFile:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self.data = data

    def getvalue(self):
        return self.data


def test_excel_upload_and_validation():
    # 1. Test loading valid excel with errors
    excel_bytes = generate_test_excel_with_errors()
    upload_file = DummyUploadFile("students_test.xlsx", excel_bytes)
    
    df, err = load_student_data(upload_file)
    assert err is None, f"Failed to load excel: {err}"
    assert df is not None
    assert len(df) == 7, f"Expected 7 rows in test file, got {len(df)}"

    # 2. Test column suggestion heuristics
    name_col = suggest_default_column(list(df.columns), ["name", "student_name", "candidate", "full_name"])
    roll_col = suggest_default_column(list(df.columns), ["roll", "roll_number", "id", "reg_no", "reg"])
    
    assert name_col == "Student Name", f"Expected 'Student Name', got {name_col}"
    assert roll_col == "Roll Number", f"Expected 'Roll Number', got {roll_col}"

    # 3. Test validation logic
    res = validate_student_records(df, name_col=name_col, roll_col=roll_col)
    
    valid_df = res["valid_df"]
    skipped_df = res["skipped_df"]

    assert res["total_rows"] == 7
    # 3 valid records: Marcus Vance (REG-1001), Elena Rostova (REG-1002), "  O'Neil MacDonald  " (00987)
    assert res["valid_count"] == 3, f"Expected 3 valid rows, got {res['valid_count']}"
    assert res["skipped_count"] == 4, f"Expected 4 skipped rows, got {res['skipped_count']}"
    assert res["empty_name_count"] == 1, f"Expected 1 empty name, got {res['empty_name_count']}"
    assert res["empty_roll_count"] == 1, f"Expected 1 empty roll, got {res['empty_roll_count']}"
    assert res["duplicate_count"] == 1, f"Expected 1 duplicate, got {res['duplicate_count']}"
    assert res["empty_row_count"] == 1, f"Expected 1 empty row, got {res['empty_row_count']}"

    # 4. Test non-destructive string preservation
    # Check that "  O'Neil MacDonald  " kept its exact leading/trailing spaces and apostrophe
    special_rec = valid_df[valid_df["Roll Number"] == "00987"].iloc[0]
    assert special_rec["Student Name"] == "  O'Neil MacDonald  ", f"Name was modified: {repr(special_rec['Student Name'])}"
    assert special_rec["Roll Number"] == "00987", f"Roll number was modified: {repr(special_rec['Roll Number'])}"

    print("[PASS] Excel upload & validation tests passed!")


def test_csv_upload_and_custom_columns():
    csv_content = """Candidate_Full_Name,Enrollment_ID,Score
Alice Smith,ENR-201,95
Bob Johnson,ENR-202,88
,ENR-203,70
Charlie Brown,,80
Alice Smith,ENR-201,95
"""
    upload_file = DummyUploadFile("participants.csv", csv_content.encode("utf-8"))
    
    df, err = load_student_data(upload_file)
    assert err is None
    assert len(df) == 5

    # Test column heuristics with different header names
    name_col = suggest_default_column(list(df.columns), ["name", "candidate", "student", "full_name"])
    roll_col = suggest_default_column(list(df.columns), ["roll", "enrollment", "id", "reg"])

    assert name_col == "Candidate_Full_Name", f"Expected Candidate_Full_Name, got {name_col}"
    assert roll_col == "Enrollment_ID", f"Expected Enrollment_ID, got {roll_col}"

    res = validate_student_records(df, name_col, roll_col)
    assert res["valid_count"] == 2
    assert res["skipped_count"] == 3
    print("[PASS] CSV upload & custom column tests passed!")


if __name__ == "__main__":
    test_excel_upload_and_validation()
    test_csv_upload_and_custom_columns()
    print("[SUCCESS] ALL TESTS PASSED SUCCESSFULLY!")
