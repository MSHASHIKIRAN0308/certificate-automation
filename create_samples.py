import os
from utils.excel_handler import generate_sample_excel, generate_test_excel_with_errors
from utils.preview_renderer import create_default_template_image
import pandas as pd
from io import BytesIO

def main():
    os.makedirs("sample_data", exist_ok=True)
    
    excel_bytes = generate_sample_excel()
    with open("sample_data/sample_students.xlsx", "wb") as f:
        f.write(excel_bytes)
    print("Created sample_data/sample_students.xlsx")

    # Also save as CSV for quick CSV testing
    df_sample = pd.read_excel(BytesIO(excel_bytes), engine="openpyxl")
    df_sample.to_csv("sample_data/sample_students.csv", index=False)
    print("Created sample_data/sample_students.csv")

    error_bytes = generate_test_excel_with_errors()
    with open("sample_data/sample_students_with_errors.xlsx", "wb") as f:
        f.write(error_bytes)
    print("Created sample_data/sample_students_with_errors.xlsx")

    img = create_default_template_image(1920, 1080)
    img.save("sample_data/sample_template.png")
    print("Created sample_data/sample_template.png")

if __name__ == "__main__":
    main()
