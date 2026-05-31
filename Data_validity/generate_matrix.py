"""
Three-Matrix Column Audit Generator
-------------------------------------
Produces a single Excel workbook with three sheets:

  Sheet 1 – Column Presence   : Is the column present in the file?
  Sheet 2 – Has Data           : Does the column contain at least one non-null, non-zero value?
  Sheet 3 – Is Blank           : Is the column entirely null/zero?

Rows    → expected column names (headers of col_names.xlsx)
Columns → one per data file, plus a "Consistent" summary column
Values  → YES / NO  (no colour grading)

Usage:
    python generate_matrix.py

Requirements:
    pip install pandas openpyxl
"""

import os
import glob
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
COL_NAMES_FILE  = "col_names.xlsx"
DATA_FILES_GLOB = "*.xlsx"
EXCLUDE_FILES   = {"col_names.xlsx", "example.xlsx", "column_matrix.xlsx"}
OUTPUT_FILE     = "column_matrix.xlsx"
# ─────────────────────────────────────────────


def get_row_names(col_names_file: str) -> list[str]:
    df = pd.read_excel(col_names_file, nrows=0)
    return list(df.columns)


def analyse_file(filepath: str, row_names: list[str]) -> dict:
    """
    Returns a dict with three sub-dicts for every expected column name:
      presence  : YES / NO
      has_data  : YES / NO / N/A  (N/A when column is absent)
      is_blank  : YES / NO / N/A  (N/A when column is absent)
    """
    df = pd.read_excel(filepath)
    file_cols = set(df.columns)

    presence = {}
    has_data = {}
    is_blank = {}

    for col in row_names:
        if col not in file_cols:
            presence[col] = "NO"
            has_data[col] = "N/A"
            is_blank[col] = "N/A"
        else:
            presence[col] = "YES"
            series = df[col].replace(0, np.nan)   # treat 0 as null
            non_null_exists = series.notna().any()
            has_data[col] = "YES" if non_null_exists else "NO"
            is_blank[col] = "YES" if not non_null_exists else "NO"

    return {"presence": presence, "has_data": has_data, "is_blank": is_blank}


def build_matrices(row_names: list[str], data_files: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Build all three DataFrames."""
    pres_records = {}
    data_records = {}
    blank_records = {}

    for fpath in data_files:
        fname = os.path.basename(fpath)
        result = analyse_file(fpath, row_names)
        pres_records[fname]  = result["presence"]
        data_records[fname]  = result["has_data"]
        blank_records[fname] = result["is_blank"]

    def make_df(records: dict) -> pd.DataFrame:
        df = pd.DataFrame(records, index=row_names)
        df.index.name = "Column Name"
        file_cols = list(df.columns)
        # Consistent = YES only when all file values are YES
        df["Consistent"] = df[file_cols].apply(
            lambda row: "YES" if all(v == "YES" for v in row) else "NO",
            axis=1
        )
        return df.reset_index()

    return make_df(pres_records), make_df(data_records), make_df(blank_records)


def style_sheet(ws, title: str) -> None:
    """Apply plain, readable formatting — no colour grading, just borders + bold header."""
    HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")   # light grey header
    HEADER_FONT = Font(bold=True, name="Arial", size=10)
    BODY_FONT   = Font(name="Arial", size=10)
    THIN        = Side(border_style="thin", color="BFBFBF")
    BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CENTER      = Alignment(horizontal="center", vertical="center")
    LEFT        = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    max_col = ws.max_column
    max_row = ws.max_row

    # Header row
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.border    = BORDER
        cell.alignment = LEFT if col_idx == 1 else CENTER

    # Body rows
    for row_idx in range(2, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = LEFT if col_idx == 1 else CENTER

    # Column widths
    ws.column_dimensions["A"].width = 70
    for col_idx in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

    # Row heights
    for row_idx in range(1, max_row + 1):
        ws.row_dimensions[row_idx].height = 18

    # Freeze header + column-name column
    ws.freeze_panes = "B2"

    # Sheet tab title
    ws.title = title


def write_workbook(
    df_pres: pd.DataFrame,
    df_data: pd.DataFrame,
    df_blank: pd.DataFrame,
    output_file: str,
) -> None:
    sheets = [
        ("1. Column Presence",  df_pres),
        ("2. Has Data",         df_data),
        ("3. Is Blank",         df_blank),
    ]

    # Write all sheets via pandas ExcelWriter
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, df in sheets:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Now open and style
    wb = load_workbook(output_file)
    for sheet_name, _ in sheets:
        style_sheet(wb[sheet_name], sheet_name)

    wb.save(output_file)
    print(f"✅  Workbook saved → {output_file}")


def main():
    row_names = get_row_names(COL_NAMES_FILE)
    print(f"Loaded {len(row_names)} expected columns from '{COL_NAMES_FILE}'")

    all_xlsx = glob.glob(DATA_FILES_GLOB)
    data_files = sorted(
        f for f in all_xlsx
        if os.path.basename(f) not in EXCLUDE_FILES
    )
    if not data_files:
        raise FileNotFoundError("No data files found. Check DATA_FILES_GLOB / EXCLUDE_FILES.")
    print(f"Found {len(data_files)} data file(s): {[os.path.basename(f) for f in data_files]}")

    df_pres, df_data, df_blank = build_matrices(row_names, data_files)

    write_workbook(df_pres, df_data, df_blank, OUTPUT_FILE)

    # Quick summary
    for label, df in [("Presence", df_pres), ("Has Data", df_data), ("Is Blank", df_blank)]:
        yes_count = (df["Consistent"] == "YES").sum()
        print(f"  [{label}] Consistent YES: {yes_count}/{len(df)}")


if __name__ == "__main__":
    main()