"""
Power BI Query & Sample Data Report
------------------------------------
Point this at a .pbix file. It will:
  1. Pull out every Power Query (M) query defined in the file.
  2. Pull the first 10 rows of the actual loaded data for each query's table
     (straight from the .pbix's embedded data model - no need to open Power BI).
  3. Write it all into a Word document saved in the SAME folder as the .pbix file.

Requirements:
    pip install pbixray python-docx pandas

Notes:
  - This reads the model that's already inside the .pbix (what you last saved/
    refreshed in Power BI Desktop). It does not re-run the queries against your
    original data sources, so "sample data" reflects the last refresh, not a
    live re-fetch.
  - Very large tables can take a moment to decode - only the first 10 rows are
    kept in the report, but pbixray decodes the queried columns before slicing.
"""

import os
from pathlib import Path

import pandas as pd

try:
    from pbixray import PBIXRay
except ImportError:
    raise SystemExit(
        "The 'pbixray' package isn't installed. Run:\n"
        "    pip install pbixray\n"
        "then try again."
    )

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor
except ImportError:
    raise SystemExit(
        "The 'python-docx' package isn't installed. Run:\n"
        "    pip install python-docx\n"
        "then try again."
    )

MAX_COLUMNS_IN_TABLE = (
    12  # keep the Word table readable; extra columns get listed below it
)
MONOSPACE_FONT = "Consolas"


# ===================== INPUT =====================


def clean_path_input(raw):
    return raw.strip().strip('"').strip("'")


def prompt_pbix_path():
    while True:
        raw = input("\nFull path to the .pbix file: ")
        path = Path(clean_path_input(raw))
        if not path.exists():
            print(f"  '{path}' doesn't exist. Please try again.")
            continue
        if path.suffix.lower() != ".pbix":
            proceed = (
                input(
                    f"  '{path.name}' doesn't end in .pbix - continue anyway? (y/n): "
                )
                .strip()
                .lower()
            )
            if proceed != "y":
                continue
        return path


# ===================== DOCX HELPERS =====================


def add_m_code_paragraph(doc, m_code):
    """Render multi-line M code as a monospace block (python-docx needs
    explicit line breaks - a literal '\\n' inside run text won't wrap)."""
    p = doc.add_paragraph()
    lines = str(m_code).splitlines() or [""]
    for i, line in enumerate(lines):
        run = p.add_run(line)
        run.font.name = MONOSPACE_FONT
        run.font.size = Pt(9)
        if i < len(lines) - 1:
            run.add_break()
    p.paragraph_format.space_after = Pt(8)
    return p


def add_dataframe_table(doc, df):
    """Insert a small Word table showing the given (already-truncated) dataframe."""
    shown_cols = list(df.columns[:MAX_COLUMNS_IN_TABLE])
    extra_cols = list(df.columns[MAX_COLUMNS_IN_TABLE:])

    table = doc.add_table(rows=1, cols=len(shown_cols))
    table.style = "Light Grid Accent 1"

    header_cells = table.rows[0].cells
    for i, col_name in enumerate(shown_cols):
        header_cells[i].text = str(col_name)
        for paragraph in header_cells[i].paragraphs:
            for run in paragraph.runs:
                run.bold = True

    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, col_name in enumerate(shown_cols):
            value = row[col_name]
            cells[i].text = "" if pd.isna(value) else str(value)

    if extra_cols:
        note = doc.add_paragraph()
        note_run = note.add_run(
            f"(+{len(extra_cols)} more column(s) not shown: {', '.join(map(str, extra_cols))})"
        )
        note_run.italic = True
        note_run.font.size = Pt(9)


# ===================== REPORT BUILD =====================


def build_report(pbix_path):
    print(f"\nOpening: {pbix_path}")
    model = PBIXRay(str(pbix_path))

    pq_df = model.power_query
    if pq_df is None or pq_df.empty:
        print("No Power Query (M) queries were found in this file.")
        pq_df = pd.DataFrame(columns=["TableName", "Expression"])
    else:
        print(f"Found {len(pq_df)} query/queries.")

    doc = Document()

    title = doc.add_heading(f"Power BI Query Report - {pbix_path.stem}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    intro = doc.add_paragraph()
    intro.add_run(f"Source file: {pbix_path.name}").italic = True
    doc.add_paragraph()  # spacer

    for i, row in pq_df.iterrows():
        table_name = row["TableName"]
        expression = row["Expression"]

        doc.add_heading(str(table_name), level=1)

        doc.add_heading("Power Query (M) Steps", level=2)
        add_m_code_paragraph(doc, expression)

        doc.add_heading("Sample Data (first 10 rows)", level=2)
        try:
            full_table = model.get_table(table_name)
            row_count = len(full_table)
            preview = full_table.head(10)
            doc.add_paragraph(f"Total rows loaded in model: {row_count:,}")
            if preview.empty:
                doc.add_paragraph("(Table has no rows.)")
            else:
                add_dataframe_table(doc, preview)
        except Exception as e:
            doc.add_paragraph(f"Could not load sample data for this table: {e}")

        if i < len(pq_df) - 1:
            doc.add_page_break()

    output_path = pbix_path.with_name(f"{pbix_path.stem}_PowerQuery_Report.docx")
    doc.save(output_path)
    return output_path


def main():
    pbix_path = prompt_pbix_path()
    output_path = build_report(pbix_path)
    print(f"\nDone! Report saved to:\n  {output_path}")


if __name__ == "__main__":
    main()
