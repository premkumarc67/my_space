"""
Three-Matrix Column Audit Generator
-------------------------------------
Produces a single Excel workbook with three sheets:

  Sheet 1 – Column Presence : Is the column present in the file?
  Sheet 2 – Has Data        : Does the column contain at least one non-null, non-zero value?
  Sheet 3 – Is Blank        : Is the column entirely null/zero?

Rows    → expected column names (headers of col_names.xlsx)
Columns → one per data file, plus a "Consistent" summary column
Values  → YES / NO / N/A

Color logic (per sheet):
  Sheet 1 – Presence  : YES=green, NO=red,    N/A=grey
  Sheet 2 – Has Data  : YES=green, NO=orange, N/A=grey
  Sheet 3 – Is Blank  : YES=red,   NO=green,  N/A=grey
  Consistent col      : YES=dark-green/bold, NO=dark-red/bold  (all sheets)
  Header row          : dark navy, white text
  Column Name col     : plain white background

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


# ─────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────
# Shared
HEADER_BG    = "1F3864"   # dark navy
HEADER_FG    = "FFFFFF"

# YES / NO / N/A fills  (background hex, no #)
GREEN_BG     = "C6EFCE"   # soft green  → positive YES
GREEN_FG     = "1E6B2E"
ORANGE_BG    = "FFE0B2"   # soft orange → warning NO (has-data sheet)
ORANGE_FG    = "7B3F00"
RED_BG       = "FFCCCC"   # soft red    → negative YES / NO
RED_FG       = "9C0006"
GREY_BG      = "F2F2F2"   # N/A
GREY_FG      = "808080"

# Consistent column (stronger, bold)
C_YES_BG     = "375623"   # deep green
C_YES_FG     = "FFFFFF"
C_NO_BG      = "C00000"   # deep red
C_NO_FG      = "FFFFFF"
C_NA_BG      = "808080"
C_NA_FG      = "FFFFFF"

ROW_ALT_BG   = "FAFAFA"   # very light stripe for readability
ROW_NORM_BG  = "FFFFFF"
# ─────────────────────────────────────────────


def _fill(hex_bg):
    return PatternFill("solid", fgColor=hex_bg)

def _font(hex_fg, bold=False):
    return Font(name="Arial", size=10, color=hex_fg, bold=bold)


# Per-sheet colour maps:
#   key = cell value → (bg, fg, bold)
SHEET_COLOURS = {
    "1. Column Presence": {
        "YES": (GREEN_BG,  GREEN_FG,  False),
        "NO":  (RED_BG,    RED_FG,    False),
        "N/A": (GREY_BG,   GREY_FG,   False),
    },
    "2. Has Data": {
        "YES": (GREEN_BG,  GREEN_FG,  False),
        "NO":  (ORANGE_BG, ORANGE_FG, False),
        "N/A": (GREY_BG,   GREY_FG,   False),
    },
    "3. Is Blank": {
        "YES": (RED_BG,    RED_FG,    False),   # blank = bad
        "NO":  (GREEN_BG,  GREEN_FG,  False),
        "N/A": (GREY_BG,   GREY_FG,   False),
    },
}

CONSISTENT_COLOURS = {
    "YES": (C_YES_BG, C_YES_FG, True),
    "NO":  (C_NO_BG,  C_NO_FG,  True),
    "N/A": (C_NA_BG,  C_NA_FG,  True),
}


def get_row_names(col_names_file: str) -> list[str]:
    df = pd.read_excel(col_names_file, nrows=0)
    return list(df.columns)


def analyse_file(filepath: str, row_names: list[str]) -> dict:
    df = pd.read_excel(filepath)
    file_cols = set(df.columns)

    presence, has_data, is_blank = {}, {}, {}

    for col in row_names:
        if col not in file_cols:
            presence[col] = "NO"
            has_data[col] = "N/A"
            is_blank[col] = "N/A"
        else:
            presence[col] = "YES"
            series = df[col].replace(0, np.nan)
            non_null_exists = series.notna().any()
            has_data[col] = "YES" if non_null_exists else "NO"
            is_blank[col] = "YES" if not non_null_exists else "NO"

    return {"presence": presence, "has_data": has_data, "is_blank": is_blank}


def build_matrices(row_names, data_files):
    pres_rec, data_rec, blank_rec = {}, {}, {}

    for fpath in data_files:
        fname  = os.path.basename(fpath)
        result = analyse_file(fpath, row_names)
        pres_rec[fname]  = result["presence"]
        data_rec[fname]  = result["has_data"]
        blank_rec[fname] = result["is_blank"]

    def make_df(records):
        df = pd.DataFrame(records, index=row_names)
        df.index.name = "Column Name"
        file_cols = list(df.columns)
        df["Consistent"] = df[file_cols].apply(
            lambda row: "YES" if all(v == "YES" for v in row) else "NO", axis=1
        )
        return df.reset_index()

    return make_df(pres_rec), make_df(data_rec), make_df(blank_rec)


def style_sheet(ws, sheet_name: str) -> None:
    colour_map   = SHEET_COLOURS[sheet_name]
    thin         = Side(border_style="thin", color="D0D0D0")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)
    center       = Alignment(horizontal="center", vertical="center")
    left_wrap    = Alignment(horizontal="left", vertical="center", wrap_text=True)

    max_col = ws.max_column
    max_row = ws.max_row

    # ── Header row ───────────────────────────────────────────
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = _fill(HEADER_BG)
        cell.font      = _font(HEADER_FG, bold=True)
        cell.border    = border
        cell.alignment = left_wrap if col_idx == 1 else center

    # ── Body rows ────────────────────────────────────────────
    for row_idx in range(2, max_row + 1):
        alt_bg = ROW_ALT_BG if row_idx % 2 == 0 else ROW_NORM_BG

        for col_idx in range(1, max_col + 1):
            cell  = ws.cell(row=row_idx, column=col_idx)
            value = str(cell.value).strip() if cell.value is not None else ""
            cell.border    = border

            if col_idx == 1:
                # Column Name — neutral alternating stripe
                cell.fill      = _fill(alt_bg)
                cell.font      = _font("000000")
                cell.alignment = left_wrap

            elif col_idx == max_col:
                # Consistent column — strong colours
                bg, fg, bold = CONSISTENT_COLOURS.get(value, (GREY_BG, GREY_FG, False))
                cell.fill      = _fill(bg)
                cell.font      = _font(fg, bold=bold)
                cell.alignment = center

            else:
                # File value columns
                bg, fg, bold = colour_map.get(value, (alt_bg, "000000", False))
                cell.fill      = _fill(bg)
                cell.font      = _font(fg, bold=bold)
                cell.alignment = center

    # ── Column widths ────────────────────────────────────────
    ws.column_dimensions["A"].width = 70
    for col_idx in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

    # ── Row heights ──────────────────────────────────────────
    ws.row_dimensions[1].height = 22
    for row_idx in range(2, max_row + 1):
        ws.row_dimensions[row_idx].height = 18

    # ── Freeze panes ─────────────────────────────────────────
    ws.freeze_panes = "B2"

    ws.title = sheet_name


def write_workbook(df_pres, df_data, df_blank, output_file):
    sheets = [
        ("1. Column Presence", df_pres),
        ("2. Has Data",        df_data),
        ("3. Is Blank",        df_blank),
    ]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, df in sheets:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(output_file)
    for sheet_name, _ in sheets:
        style_sheet(wb[sheet_name], sheet_name)

    wb.save(output_file)
    print(f"✅  Workbook saved → {output_file}")


def main():
    row_names = get_row_names(COL_NAMES_FILE)
    print(f"Loaded {len(row_names)} expected columns from '{COL_NAMES_FILE}'")

    all_xlsx   = glob.glob(DATA_FILES_GLOB)
    data_files = sorted(f for f in all_xlsx if os.path.basename(f) not in EXCLUDE_FILES)

    if not data_files:
        raise FileNotFoundError("No data files found. Check DATA_FILES_GLOB / EXCLUDE_FILES.")
    print(f"Found {len(data_files)} data file(s): {[os.path.basename(f) for f in data_files]}")

    df_pres, df_data, df_blank = build_matrices(row_names, data_files)
    write_workbook(df_pres, df_data, df_blank, OUTPUT_FILE)

    for label, df in [("Presence", df_pres), ("Has Data", df_data), ("Is Blank", df_blank)]:
        yes_count = (df["Consistent"] == "YES").sum()
        print(f"  [{label}] Consistent YES: {yes_count}/{len(df)}")


if __name__ == "__main__":
    main()