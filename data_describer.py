"""
Dataset Summary Report Generator
---------------------------------
Scans a folder (and all its sub-folders) for CSV, Excel (.xlsx/.xls) and TXT
data files, and builds a single Word (.docx) report containing, for every
file found:

    - One table per file
    - One row per column in that dataset, showing:
        Column Name | Data Type | Row 1 | Row 2 | Row 3 | Row 4 |
        Mean | Median | Mode | Range | Outliers

Mean / Median / Range / Outliers are computed only for numeric columns
(computed over the FULL column, not just the first 4 rows, so they're
statistically meaningful). Mode is computed for every column type.

Requirements (install once):
    pip install pandas python-docx openpyxl

Run it, then answer the two prompts it shows you.
"""

import os
import sys
import time
import statistics
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Font size used for every table in the report
TABLE_FONT_SIZE = Pt(10)

# File extensions we treat as "datasets"
DATA_EXTENSIONS = {".csv", ".xlsx", ".xls", ".txt"}


# --------------------------------------------------------------------------
# Reading data files
# --------------------------------------------------------------------------
def read_dataset(file_path: Path):
    """
    Reads a CSV / Excel / TXT file into a pandas DataFrame.
    Returns None (and prints a warning) if the file can't be read.
    """
    ext = file_path.suffix.lower()
    try:
        if ext == ".csv":
            return pd.read_csv(file_path, encoding="utf-8", on_bad_lines="skip")
        elif ext in (".xlsx", ".xls"):
            return pd.read_excel(file_path)
        elif ext == ".txt":
            # Auto-detect the delimiter (comma, tab, pipe, semicolon, whitespace...)
            return pd.read_csv(
                file_path,
                sep=None,
                engine="python",
                encoding="utf-8",
                on_bad_lines="skip",
            )
    except UnicodeDecodeError:
        # Retry with a more forgiving encoding
        try:
            if ext == ".csv":
                return pd.read_csv(file_path, encoding="latin-1", on_bad_lines="skip")
            elif ext == ".txt":
                return pd.read_csv(
                    file_path,
                    sep=None,
                    engine="python",
                    encoding="latin-1",
                    on_bad_lines="skip",
                )
        except Exception as e:
            print(f"  [skipped] Could not read {file_path.name}: {e}")
            return None
    except Exception as e:
        print(f"  [skipped] Could not read {file_path.name}: {e}")
        return None


# --------------------------------------------------------------------------
# Per-column statistics
# --------------------------------------------------------------------------
def get_mode(series: pd.Series):
    try:
        modes = series.mode(dropna=True)
        if modes.empty:
            return "N/A"
        # Show up to 3 mode values if there are several
        vals = [str(v) for v in modes.tolist()[:3]]
        return ", ".join(vals) + (" ..." if len(modes) > 3 else "")
    except Exception:
        return "N/A"


def get_outliers(series: pd.Series):
    """IQR method: values outside [Q1 - 1.5*IQR, Q3 + 1.5*IQR]."""
    try:
        clean = series.dropna()
        if len(clean) < 4:
            return "N/A"
        q1 = clean.quantile(0.25)
        q3 = clean.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            return "None"
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        outliers = clean[(clean < lower) | (clean > upper)]
        if outliers.empty:
            return "None"
        count = len(outliers)
        sample = ", ".join(
            str(round(v, 2)) if isinstance(v, float) else str(v)
            for v in outliers.head(3).tolist()
        )
        return f"{count} found (e.g. {sample})"
    except Exception:
        return "N/A"


def column_stats(series: pd.Series):
    """Returns a dict with data_type, mean, median, mode, range, outliers."""
    dtype = str(series.dtype)
    # Booleans register as "numeric" in pandas but max()-min() on bools
    # raises TypeError, and mean/median of True/False isn't meaningful here
    # -> treat bool columns like categorical ones (mode only).
    is_numeric = pd.api.types.is_numeric_dtype(
        series
    ) and not pd.api.types.is_bool_dtype(series)

    stats = {
        "data_type": dtype,
        "mean": "N/A",
        "median": "N/A",
        "mode": get_mode(series),
        "range": "N/A",
        "outliers": "N/A",
    }

    if is_numeric:
        clean = series.dropna()
        if not clean.empty:
            try:
                clean = clean.astype(float)  # normalize any odd numeric subtype
                stats["mean"] = round(clean.mean(), 3)
                stats["median"] = round(clean.median(), 3)
                stats["range"] = round(clean.max() - clean.min(), 3)
                stats["outliers"] = get_outliers(clean)
            except Exception as e:
                stats["mean"] = stats["median"] = stats["range"] = stats["outliers"] = (
                    f"N/A ({e})"
                )

    return stats


# --------------------------------------------------------------------------
# Finding data files
# --------------------------------------------------------------------------
def find_data_files(root_folder: Path):
    """Recursively finds all CSV/Excel/TXT files under root_folder."""
    found = []
    for dirpath, _dirnames, filenames in os.walk(root_folder):
        for fname in filenames:
            if Path(fname).suffix.lower() in DATA_EXTENSIONS:
                found.append(Path(dirpath) / fname)
    return sorted(found)


# --------------------------------------------------------------------------
# Word report building
# --------------------------------------------------------------------------
def set_table_font(table, size=TABLE_FONT_SIZE):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = size


def add_file_table(document: Document, file_path: Path, df: pd.DataFrame):
    # File name heading
    heading = document.add_heading(file_path.name, level=2)
    for run in heading.runs:
        run.font.size = Pt(13)

    path_para = document.add_paragraph(f"Path: {file_path}")
    path_para.runs[0].font.size = Pt(9)
    document.add_paragraph(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns").runs[
        0
    ].font.size = Pt(9)

    headers = [
        "Column Name",
        "Data Type",
        "Row 1",
        "Row 2",
        "Row 3",
        "Row 4",
        "Mean",
        "Median",
        "Mode",
        "Range",
        "Outliers",
    ]
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Light Grid Accent 1"

    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True

    # First 4 data rows (excluding header, which pandas already treats as header)
    first_rows = df.head(4)

    for col in df.columns:
        try:
            stats = column_stats(df[col])
        except Exception as e:
            stats = {
                "data_type": str(df[col].dtype),
                "mean": "N/A",
                "median": "N/A",
                "mode": "N/A",
                "range": "N/A",
                "outliers": f"N/A ({e})",
            }
        row_cells = table.add_row().cells

        row_values = []
        for i in range(4):
            if i < len(first_rows):
                val = first_rows.iloc[i][col]
                row_values.append("" if pd.isna(val) else str(val))
            else:
                row_values.append("")

        values = [
            str(col),
            stats["data_type"],
            *row_values,
            str(stats["mean"]),
            str(stats["median"]),
            str(stats["mode"]),
            str(stats["range"]),
            str(stats["outliers"]),
        ]
        for i, v in enumerate(values):
            row_cells[i].text = v

    set_table_font(table)
    document.add_paragraph("")  # spacing after table


def build_report(root_folder: Path, output_path: Path, max_workers: int = None):
    data_files = find_data_files(root_folder)

    if not data_files:
        print("No CSV, Excel or TXT files found in that folder (or its sub-folders).")
        return

    if max_workers is None:
        # Use all CPU cores, but no point spinning up more workers than files
        max_workers = min(len(data_files), os.cpu_count() or 4)

    print(
        f"Found {len(data_files)} data file(s). Reading {max_workers} at a time in parallel..."
    )
    start = time.time()

    # --- Read every file's DataFrame in parallel (this is the slow, CPU-bound part) ---
    results = {}  # file_path -> DataFrame (or None if it failed)
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_path = {
            executor.submit(read_dataset, file_path): file_path
            for file_path in data_files
        }
        done_count = 0
        for future in as_completed(future_to_path):
            file_path = future_to_path[future]
            done_count += 1
            try:
                df = future.result()
            except Exception as e:
                print(
                    f"  [{done_count}/{len(data_files)}] FAILED: {file_path.name} ({e})"
                )
                df = None
            else:
                status = "ok" if df is not None and not df.empty else "empty/skipped"
                print(f"  [{done_count}/{len(data_files)}] {status}: {file_path.name}")
            results[file_path] = df

    elapsed_read = time.time() - start
    print(
        f"Finished reading all files in {elapsed_read:.1f}s. Building the Word document..."
    )

    # --- Build the .docx sequentially (fast — no I/O), in the original sorted order ---
    document = Document()
    document.add_heading("Dataset Summary Report", level=1)
    document.add_paragraph(f"Source folder: {root_folder}")
    document.add_paragraph(f"Files processed: {len(data_files)}")
    document.add_paragraph("")

    processed = 0
    for (
        file_path
    ) in data_files:  # data_files is already sorted, so output order is stable
        df = results.get(file_path)
        if df is None or df.empty:
            continue
        add_file_table(document, file_path, df)
        processed += 1

    document.save(output_path)
    total_elapsed = time.time() - start
    print(
        f"\nDone in {total_elapsed:.1f}s. Processed {processed}/{len(data_files)} file(s)."
    )
    print(f"Report saved to: {output_path}")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def main():
    folder_input = input("Enter the folder location to scan: ").strip().strip('"')
    root_folder = Path(folder_input)

    if not root_folder.exists() or not root_folder.is_dir():
        print(f"Error: '{root_folder}' is not a valid folder.")
        sys.exit(1)

    output_input = (
        input(
            "Enter the full path (including file name) to save the report "
            "(e.g. C:\\reports\\summary.docx): "
        )
        .strip()
        .strip('"')
    )
    output_path = Path(output_input)

    if output_path.suffix.lower() != ".docx":
        output_path = output_path.with_suffix(".docx")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    build_report(root_folder, output_path)


if __name__ == "__main__":
    main()
