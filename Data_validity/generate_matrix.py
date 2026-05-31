"""
Column Presence Matrix Generator
---------------------------------
Reads multiple Excel files, checks which columns (from col_names.xlsx)
are present in each file, and outputs a matrix showing YES/blank per file.
Also adds a "Consistent" column indicating if the column is present across ALL files.

Usage:
    python generate_matrix.py

Requirements:
    pip install pandas openpyxl
"""

import os
import glob
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# CONFIG — update these paths as needed
# ─────────────────────────────────────────────
COL_NAMES_FILE = "col_names.xlsx"       # File whose column headers are the row names
DATA_FILES_GLOB = "*.xlsx"              # Glob pattern to pick up all data files
EXCLUDE_FILES   = {"col_names.xlsx", "example.xlsx", "generate_matrix.py"}
OUTPUT_FILE     = "column_matrix.xlsx"
# ─────────────────────────────────────────────


def get_row_names(col_names_file: str) -> list[str]:
    """Return the ordered list of expected column names from col_names.xlsx."""
    df = pd.read_excel(col_names_file, nrows=0)   # header only, no data rows needed
    return list(df.columns)


def get_file_columns(filepath: str) -> set[str]:
    """Return the set of column names present in the first sheet of an Excel file."""
    df = pd.read_excel(filepath, nrows=0)
    return set(df.columns)


def build_matrix(row_names: list[str], data_files: list[str]) -> pd.DataFrame:
    """
    Build the presence matrix.

    Rows    → expected column names (from col_names.xlsx)
    Columns → data file names (filename only, no path)
    Values  → "YES" if the column exists in that file, else "" (blank)
    Last col→ "Consistent" = "YES" if present in ALL files, else "NO"
    """
    records = {}
    for fpath in data_files:
        fname = os.path.basename(fpath)
        present = get_file_columns(fpath)
        records[fname] = {col: "YES" if col in present else "" for col in row_names}

    df = pd.DataFrame(records, index=row_names)
    df.index.name = "Column Name"

    file_cols = list(df.columns)
    df["Consistent"] = df[file_cols].apply(
        lambda row: "YES" if all(v == "YES" for v in row) else "NO",
        axis=1
    )
    return df.reset_index()


def style_workbook(output_file: str, file_cols: list[str]) -> None:
    """Apply formatting: header row, YES/NO colours, column widths."""
    wb = load_workbook(output_file)
    ws = wb.active

    # Colour palette
    HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")   # dark blue
    HEADER_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    YES_FILL      = PatternFill("solid", fgColor="C6EFCE")   # light green
    YES_FONT      = Font(color="276221", name="Arial", size=10)
    NO_FILL       = PatternFill("solid", fgColor="FFCCCC")   # light red
    NO_FONT       = Font(color="9C0006", name="Arial", size=10)
    BLANK_FILL    = PatternFill("solid", fgColor="FFF2CC")   # light yellow
    BODY_FONT     = Font(name="Arial", size=10)
    THIN          = Side(border_style="thin", color="D3D3D3")
    BORDER        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CENTER        = Alignment(horizontal="center", vertical="center", wrap_text=False)
    LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    max_col = ws.max_column
    max_row = ws.max_row

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.border = BORDER
        cell.alignment = CENTER if col_idx > 1 else LEFT

    for row_idx in range(2, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell  = ws.cell(row=row_idx, column=col_idx)
            value = cell.value
            cell.border = BORDER

            if col_idx == 1:                                  # Column Name
                cell.font      = BODY_FONT
                cell.alignment = LEFT
            elif col_idx == max_col:                          # Consistent column
                if value == "YES":
                    cell.fill, cell.font = YES_FILL, YES_FONT
                else:
                    cell.fill, cell.font = NO_FILL,  NO_FONT
                cell.alignment = CENTER
            else:                                             # File columns
                if value == "YES":
                    cell.fill, cell.font = YES_FILL,  YES_FONT
                elif value == "" or value is None:
                    cell.fill, cell.font = BLANK_FILL, BODY_FONT
                cell.alignment = CENTER

    # Column widths
    ws.column_dimensions["A"].width = 70
    for col_idx in range(2, max_col + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 22

    # Freeze header row + Column Name column
    ws.freeze_panes = "B2"

    # Auto-fit row heights
    for row_idx in range(1, max_row + 1):
        ws.row_dimensions[row_idx].height = 18

    wb.save(output_file)
    print(f"✅  Styled workbook saved → {output_file}")


def main():
    # 1. Collect row names
    row_names = get_row_names(COL_NAMES_FILE)
    print(f"Loaded {len(row_names)} expected columns from '{COL_NAMES_FILE}'")

    # 2. Find data files
    all_xlsx = glob.glob(DATA_FILES_GLOB)
    data_files = sorted(
        f for f in all_xlsx
        if os.path.basename(f) not in EXCLUDE_FILES
    )
    if not data_files:
        raise FileNotFoundError("No data files found. Check DATA_FILES_GLOB / EXCLUDE_FILES.")
    print(f"Found {len(data_files)} data file(s): {[os.path.basename(f) for f in data_files]}")

    # 3. Build matrix
    df = build_matrix(row_names, data_files)
    file_cols = [c for c in df.columns if c not in ("Column Name", "Consistent")]

    # 4. Save raw matrix
    df.to_excel(OUTPUT_FILE, index=False, sheet_name="Sheet1")
    print(f"Matrix written → {OUTPUT_FILE}")

    # 5. Apply styling
    style_workbook(OUTPUT_FILE, file_cols)

    # 6. Summary
    consistent_count = (df["Consistent"] == "YES").sum()
    print(f"\nSummary: {consistent_count}/{len(df)} columns are present in ALL files.")


if __name__ == "__main__":
    main()
