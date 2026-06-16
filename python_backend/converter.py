import os
import pandas as pd


def convert_to_xlsx(file_path, output_dir):
    """
    Converts .xls, .xlsm, .xlsx input to a clean .xlsx in output_dir.
    If already .xlsx, just copies it through (re-saves to normalize).
    Preserves all sheets.
    Returns absolute path to converted file.
    """
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")

    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xls":
        # xls needs xlrd engine
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="xlrd")
    elif ext in (".xlsm", ".xlsx"):
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_path