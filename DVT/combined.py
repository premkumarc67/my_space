import argparse
import csv
import os
import re
import sys
import tempfile
import fitz  # PyMuPDF
import cv2
import numpy as np
import pdfplumber

# ── Default column edges in PDF points (pt = 1/72 inch) ─────────────────────
# These match the standard IEC 60617 table layout (pages 1-4).
# The script now auto-detects edges per page, so these are only used when
# auto-detection produces no usable result.
DEFAULT_COL_EDGES = [54.48, 258.0, 438.0, 521.16]
#                    │      │      │      └─ right page border
#                    │      │      └──────── COMMENTS column left edge
#                    │      └─────────────── IEC DESCRIPTION column left edge
#                    └────────────────────── IEC SYMBOL column left edge

DPI   = 300           # PDF pages are rendered as images at 300 DPI. Higher Dots Per Inch - sharper image.
SCALE = DPI / 72      # PDF points → pixels conversion factor. 1 inch = 72 PDF points.
PAD   = 5             # pixels to inset from each cell border line

# ── Helpers ──────────────────────────────────────────────────────────────────

def pt_to_px(value: float) -> int:
    """Convert a PDF-point coordinate to a pixel coordinate at current DPI."""
    return int(round(value * SCALE))


def sanitize_filename(text: str, max_len: int = 60) -> str:
    """
    Turn arbitrary description text into a safe filename fragment.
    Keeps letters, digits, hyphens; collapses whitespace to underscores.
    """
    cleaned = re.sub(r"[^\w\s-]", "", text)
    cleaned = re.sub(r"\s+", "_", cleaned.strip())
    cleaned = cleaned[:max_len].strip("_")
    return cleaned or "unnamed"


CLASSES = [
    "Conductors", "Cables", "Terminals", "Plugs", "Sockets", "Switches",
    "Disconnectors", "Circuit Breakers", "Fuses", "Actuators", "Resistors",
    "Capacitors", "Inductors", "Diodes", "Thyristors", "Transistors",
    "Relays", "Contactors", "Starters", "Transformers", "Motors",
    "Generators", "Converters", "Rectifiers", "Inverters", "Batteries",
    "Meters", "Lamps", "Indicators", "Bells", "Sirens", "Pumps",
    "Fans", "Antennas", "Selectors", "Thermocouples"
]

CLASS_MAPPING = {
    "Conductors": ["conductor", "connection", "line", "wiring", "path", "neutral", "earth", "ground", "bonding", "circuit", "link", "polarity", "phase", "wire", "cable sealing", "current", "voltage", "fault", "flashover"],
    "Cables": ["cable", "coaxial"],
    "Terminals": ["terminal", "junction"],
    "Plugs": ["plug"],
    "Sockets": ["socket", "outlet"],
    "Switches": ["switch", "contact", "push-button", "dimmer"],
    "Disconnectors": ["disconnector", "isolator"],
    "Circuit Breakers": ["circuit breaker"],
    "Fuses": ["fuse"],
    "Actuators": ["actuator", "motion", "detent", "brake", "gearing", "lever", "pedal", "handwheel", "cam", "roller", "crank", "return", "action"],
    "Resistors": ["resistor", "potentiometer", "varistor", "heating element"],
    "Capacitors": ["capacitor", "capasitor"],
    "Inductors": ["inductor", "coil", "winding", "choke", "variometer"],
    "Diodes": ["diode", "photodiode"],
    "Thyristors": ["thyristor", "diac", "triac"],
    "Transistors": ["transistor", "igbt", "phototransistor"],
    "Relays": ["relay"],
    "Contactors": ["contactor"],
    "Starters": ["starter"],
    "Transformers": ["transformer"],
    "Motors": ["motor", "machine"],
    "Generators": ["generator", "primary cell", "solar", "wind", "station", "cell", "photovoltaic"],
    "Converters": ["converter", "chopper"],
    "Rectifiers": ["rectifier"],
    "Inverters": ["inverter"],
    "Batteries": ["battery"],
    "Meters": ["meter", "voltmeter", "watt-hour", "clock", "instrument"],
    "Lamps": ["lamp", "luminaire", "light", "projector"],
    "Indicators": ["indicator", "alarm"],
    "Bells": ["bell", "buzzer"],
    "Sirens": ["siren"],
    "Pumps": ["pump"],
    "Fans": ["fan"],
    "Antennas": ["antenna"],
    "Selectors": ["selector"],
    "Thermocouples": ["thermocouple", "thermal", "thrmocouple", "heat"]
}

def identify_class(desc: str, comment: str) -> str:
    text = f"{desc} {comment}".lower()
    
    # Priority multi-word/exact substring matches
    if "circuit breaker" in text: return "Circuit Breakers"
    if "disconnector" in text or "isolator" in text: return "Disconnectors"
    if "transformer" in text: return "Transformers"
    
    for cls in CLASSES:
        keywords = CLASS_MAPPING.get(cls, [cls.lower(), cls[:-1].lower() if cls.endswith('s') else cls.lower()])
        for kw in keywords:
            if re.search(r'\b' + re.escape(kw) + r'(?:s|es)?\b', text):
                return cls
                
    return "Others"


def render_pdf_pages(pdf_path: str, dpi: int, tmp_dir: str) -> list[str]:
    """
    Render every page of PDF to PNG image in tmp_dir.
    Returns a sorted list of PNG file paths, one per page.
    """
    prefix = os.path.join(tmp_dir, "page")
    doc   = fitz.open(pdf_path)
    mat   = fitz.Matrix(dpi / 72, dpi / 72) # scaling matrix - controls rendering resolution
    paths = []
    for i, page in enumerate(doc):
        pix  = page.get_pixmap(matrix=mat, alpha=False) # Converts a PDF page into an image
        path = f"{prefix}-{i + 1}.png"
        pix.save(path)
        paths.append(path)
    print(f"  Renderer: PyMuPDF  ({len(paths)} page(s))")
    return paths


def build_table_settings(col_edges: list[float]) -> dict:
    return {
        "vertical_strategy":       "explicit",
        "horizontal_strategy":     "lines",
        "explicit_vertical_lines": col_edges,
        "snap_tolerance":          3,
        "join_tolerance":          3,
        "edge_min_length":         3,
    }


def detect_col_edges_for_page(page) -> list[float] | None:
    """
    Auto-detect the three column edges for this page by finding the table
    using pdfplumber's default strategy, then extracting the x0/x1 boundaries
    of the non-None header cells.

    Returns a 4-element list [left, col1, col2, right] on success, else None.
    """
    tables = page.find_tables()
    if not tables:
        return None
    t = tables[0]
    if not t.rows:
        return None

    # Collect the x0 of every non-None cell in the header row, plus the x1
    # of the last non-None cell.  Cells whose bbox is None are merged/spanned
    # cells — we skip them.
    header_cells = [c for c in t.rows[0].cells if c is not None]
    if len(header_cells) < 3:
        return None

    left  = header_cells[0][0]   # x0 of first cell - each cwll format (x0, top, x1, bottom)
    col1  = header_cells[1][0]   # x0 of second cell (IEC DESCRIPTION)
    col2  = header_cells[2][0]   # x0 of third cell  (COMMENTS)
    right = header_cells[2][2]   # x1 of third cell  (right border) 

    return [left, col1, col2, right]


def get_cell_text(text_row: list, col_idx: int) -> str:
    """
    Safely extract text from a cell, returning '' if the index is out of range
    or the value is None.
    """
    if col_idx >= len(text_row):
        return ""
    val = text_row[col_idx]
    return (val or "").replace("\n", " ").strip()


def find_symbol_col_index(text_row: list) -> int:
    """
    In rows produced by auto-detected tables, the first non-None cell is the
    IEC SYMBOL column.  Because auto-detection may insert extra None-filled
    merged columns, we scan for the first non-None slot.  Returns 0 as default.
    """
    for i, v in enumerate(text_row):
        if v is not None:
            return i
    return 0


def find_data_col_indices(table) -> tuple[int, int, int]:
    """
    Determine which column indices correspond to (symbol, description, comments)
    by examining the header row.

    Returns (sym_idx, desc_idx, cmt_idx).
    """
    if not table.rows:
        return 0, 1, 2
    header = table.extract()[0]

    sym_idx  = 0
    desc_idx = 1
    cmt_idx  = 2

    # Find columns by matching header text
    for i, cell in enumerate(header):
        if cell is None:
            continue
        txt = cell.strip().upper()
        if "SYMBOL" in txt:
            sym_idx = i
        elif "DESCRIPTION" in txt:
            desc_idx = i
        elif "COMMENT" in txt:
            cmt_idx = i

    return sym_idx, desc_idx, cmt_idx


# ── Core extraction ───────────────────────────────────────────────────────────

def process_page(
    page_num: int,
    page,              # pdfplumber page
    img: np.ndarray,   # rendered image for this page (BGR)
    fallback_col_edges: list[float],
    output_dir: str,
) -> list[list]:
    """
    Extract all symbol rows from one page.
    Returns list of CSV rows: [page, row, description, comments, filename]

    Strategy:
      1. Auto-detect column edges from the page's own table structure.
      2. Use those edges with explicit vertical strategy so every cell is
         correctly bounded.
      3. Fall back to the user-supplied edges only if auto-detection fails.
    """
    # ── Step 1: auto-detect edges for this page ──────────────────────────────
    detected_edges = detect_col_edges_for_page(page)

    if detected_edges:
        col_edges = detected_edges
    else:
        col_edges = fallback_col_edges
        print(f"    ⚠  Could not auto-detect edges on page {page_num}. "
              f"Using fallback: {col_edges}")

    # ── Step 2: find table with the (possibly per-page) explicit edges ────────
    table_settings = build_table_settings(col_edges) # dictionary will be returned
    tables = page.find_tables(table_settings)

    if not tables:
        print(f"    ✗  No table found on page {page_num} — skipping.")
        return []

    table = tables[0]
    text_rows = table.extract() # Extract table text - [["IEC SYMBOL", "DESCRIPTION", "COMMENTS"],["", "resistor", "fixed"]]

    # Determine which column index maps to symbol / description / comments
    sym_idx, desc_idx, cmt_idx = find_data_col_indices(table)

    csv_rows = []

    for row_idx, (row, text_row) in enumerate(zip(table.rows, text_rows)):
        # Row 0 is the header ("IEC SYMBOL | IEC DESCRIPTION | COMMENTS")
        if row_idx == 0:
            continue

        # ── Get symbol cell bounding box ──────────────────────────────────────
        # row.cells - [(54.0, 100.0, 200.0, 140.0),(200.0, 100.0, 300.0, 140.0)]
        # The symbol column is always the first non-None cell in the row.
        # With explicit edges it is always cells[0].

        sym_cell = row.cells[sym_idx] if sym_idx < len(row.cells) else None
        if sym_cell is None:
            # Try cells[0] as last resort
            sym_cell = row.cells[0] if row.cells else None
        if sym_cell is None:
            print(f"    Row {row_idx}: symbol cell is None — skipping")
            continue

        x0_pdf, top_pdf, x1_pdf, bot_pdf = sym_cell 

        # ── Crop symbol from rendered image ───────────────────────────────────
        y1 = pt_to_px(top_pdf) + PAD
        y2 = pt_to_px(bot_pdf) - PAD
        x1 = pt_to_px(x0_pdf) + PAD
        x2 = pt_to_px(x1_pdf) - PAD

        crop = img[y1:y2, x1:x2]
        if crop.size == 0 or (y2 - y1) < 10:
            print(f"    Row {row_idx}: crop too small — skipping")
            continue

        # ── Extract text ──────────────────────────────────────────────────────
        desc    = get_cell_text(text_row, desc_idx)
        comment = get_cell_text(text_row, cmt_idx)

        safe  = sanitize_filename(desc)
        comp_class = identify_class(desc, comment)
        fname = f"{comp_class}_{safe}.png"
        fpath = os.path.join(output_dir, fname)

        cv2.imwrite(fpath, crop) # Save image
        csv_rows.append([page_num, row_idx, desc, comment, fname]) # Add row to CSV data
        print(f"    ✓  {fname}  ({crop.shape[1]}×{crop.shape[0]} px)")

    return csv_rows


def extract_all(
    pdf_path:   str,
    output_dir: str,
    csv_path:   str,
    col_edges:  list[float],
) -> int:
    os.makedirs(output_dir, exist_ok=True)
    all_csv_rows = []

    with tempfile.TemporaryDirectory() as tmp_dir:
        print(f"\n[1/3] Rendering pages from '{pdf_path}' at {DPI} DPI …")
        page_image_paths = render_pdf_pages(pdf_path, DPI, tmp_dir)

        print(f"\n[2/3] Extracting symbols and text …")
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) != len(page_image_paths):
                raise RuntimeError(
                    f"Page count mismatch: PDF has {len(pdf.pages)} pages "
                    f"but renderer produced {len(page_image_paths)} images."
                )

            for page_num, (page, img_path) in enumerate(zip(pdf.pages, page_image_paths), start=1):
                print(f"\n  Page {page_num}/{len(pdf.pages)}")
                img = cv2.imread(img_path)
                if img is None:
                    print(f"    ✗  Could not read rendered image: {img_path}")
                    continue

                rows = process_page(page_num, page, img, col_edges, output_dir)
                all_csv_rows.extend(rows)

    print(f"\n[3/3] Writing CSV → '{csv_path}' …")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Page", "Row", "IEC_DESCRIPTION", "COMMENTS", "Symbol_Image"])
        writer.writerows(all_csv_rows)

    return len(all_csv_rows)


# ── Diagnostic mode ───────────────────────────────────────────────────────────

def diagnose(pdf_path: str, page_index: int = 0):
    """Print word x-positions and detected edges to help re-tune for a new PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        page  = pdf.pages[page_index]
        words = page.extract_words()

    print(f"\nPage size: {page.width} × {page.height} pt")
    print(f"{'x0':>8}  {'x1':>8}  {'top':>8}  text")
    print("─" * 55)
    for w in words[:50]:
        print(f"{w['x0']:8.1f}  {w['x1']:8.1f}  {w['top']:8.1f}  {w['text']}")

    detected = detect_col_edges_for_page(page)
    print()
    if detected:
        print(f"Auto-detected edges for this page: {detected}")
    else:
        print("Could not auto-detect edges. Set --edges manually.")
    print("E.g. if columns start at x0=55, 258, 438 and end at x1=521:")
    print("  --edges 55 258 438 521")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_pdf = os.path.join(script_dir, "IEC.pdf")
    default_out = os.path.join(script_dir, "iec_symbols")
    default_csv = os.path.join(script_dir, "iec_symbols.csv")

    parser = argparse.ArgumentParser(
        description="Extract IEC 60617 symbol images + text table from a PDF"
    )
    parser.add_argument("--pdf",      default=default_pdf,
                        help="Path to the input PDF (default: IEC.pdf)")
    parser.add_argument("--out",      default=default_out,
                        help="Output directory for cropped images (default: iec_symbols/)")
    parser.add_argument("--csv",      default=default_csv,
                        help="Output CSV file path (default: iec_symbols.csv)")
    parser.add_argument("--edges",    nargs=4, type=float,
                        metavar=("LEFT", "COL1", "COL2", "RIGHT"),
                        default=DEFAULT_COL_EDGES,
                        help="Fallback column edge x-positions in PDF points used "
                             "only when auto-detection fails (default: %(default)s)")
    parser.add_argument("--diagnose", action="store_true",
                        help="Print word positions and auto-detected edges, then exit")
    parser.add_argument("--page",     type=int, default=0,
                        help="Page index for --diagnose (0-based, default: 0)")
    args = parser.parse_args()

    if args.diagnose:
        diagnose(args.pdf, args.page)
        return

    try:
        n = extract_all(
            pdf_path   = args.pdf,
            output_dir = args.out,
            csv_path   = args.csv,
            col_edges  = args.edges,
        )
        print(f"\n✅  Done — {n} symbols saved to '{args.out}/', CSV → '{args.csv}'")
    except Exception as e:
        print(f"\n❌  {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()