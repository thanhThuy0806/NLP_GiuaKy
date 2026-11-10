from pathlib import Path

import pandas as pd


ROOT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = ROOT_DIR / "CV05_CPR_03"
OUTPUT_DIR = PROJECT_DIR / "04_dong_hang"
WORK_CODE = PROJECT_DIR.name.partition("_")[2] or PROJECT_DIR.name


def convert_csv_to_xlsx_and_tsv(csv_file_path):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    base_name = OUTPUT_DIR / f"{WORK_CODE}_parallel"

    xlsx_file_path = base_name.with_suffix(".xlsx")
    tsv_file_path = base_name.with_suffix(".tsv")

    df = pd.read_csv(csv_file_path, encoding="utf-8")

    df.to_excel(xlsx_file_path, index=False, engine="openpyxl")
    print(f"✓ Đã chuyển đổi thành công sang Excel: {xlsx_file_path}")

    df.to_csv(tsv_file_path, sep="\t", index=False, encoding="utf-8")
    print(f"✓ Đã chuyển đổi thành công sang TSV: {tsv_file_path}")


if __name__ == "__main__":
    csv_file_path = ROOT_DIR / "Demo_CV05_CPR_03(Sheet1).csv"
    convert_csv_to_xlsx_and_tsv(csv_file_path)
