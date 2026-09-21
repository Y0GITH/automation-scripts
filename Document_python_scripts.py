"""
document_python_scripts.py

Recursively scans a folder for .py files, runs each one, and writes a Word
document ("Python Scripts Documentation.docx") containing, for every script:
    - the folder it was found in
    - the file name
    - the full source code
    - the captured output (stdout/stderr) from running it

Usage:
    python document_python_scripts.py
    (it will prompt you for the folder to scan and the folder to save the doc in)

Requirements:
    pip install python-docx
"""

import os
import subprocess
import sys

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("The 'python-docx' package is required. Install it with:")
    print("    pip install python-docx")
    sys.exit(1)

TIMEOUT_SECONDS = 30  # max time to let each script run before giving up on it


def find_python_files(root_folder):
    """Walk root_folder recursively and yield (folder_path, file_name) for every .py file."""
    for dirpath, _dirnames, filenames in os.walk(root_folder):
        for fname in sorted(filenames):
            if fname.lower().endswith(".py"):
                # Don't try to run this documentation script if it's sitting inside the folder
                if os.path.abspath(os.path.join(dirpath, fname)) == os.path.abspath(
                    __file__
                ):
                    continue
                yield dirpath, fname


def run_script(full_path):
    """
    Run a python script and capture its output.
    Returns a tuple: (status, output_text)
    status is one of: "OK", "ERROR", "TIMEOUT"
    """
    try:
        result = subprocess.run(
            [sys.executable, full_path],
            capture_output=True,
            text=True,
            timeout=TIMEOUT_SECONDS,
            stdin=subprocess.DEVNULL,  # scripts waiting on input() won't hang; they'll error instead
        )
        if result.returncode == 0:
            output = result.stdout if result.stdout.strip() else "(no output printed)"
            return "OK", output
        else:
            combined = (result.stdout or "") + "\n" + (result.stderr or "")
            return "ERROR", combined.strip()
    except subprocess.TimeoutExpired:
        return (
            "TIMEOUT",
            f"Script did not finish within {TIMEOUT_SECONDS} seconds and was skipped.",
        )
    except Exception as e:
        return "ERROR", f"Failed to run script: {e}"


def add_code_block(document, text):
    """Add a monospaced, shaded-looking block of text to the document."""
    para = document.add_paragraph()
    run = para.add_run(text if text.strip() else "(empty)")
    run.font.name = "Consolas"
    run.font.size = Pt(9)
    para.paragraph_format.space_after = Pt(10)


def main():
    root_folder = (
        input("Enter the folder path to scan for Python scripts: ").strip().strip('"')
    )
    if not os.path.isdir(root_folder):
        print(f"'{root_folder}' is not a valid folder.")
        sys.exit(1)

    save_folder = (
        input("Enter the folder path to save the Word document in: ").strip().strip('"')
    )
    if not os.path.isdir(save_folder):
        print(f"'{save_folder}' is not a valid folder.")
        sys.exit(1)

    files = list(find_python_files(root_folder))
    if not files:
        print("No Python files found in that folder.")
        sys.exit(0)

    print(f"Found {len(files)} Python file(s). Running each one now...\n")

    document = Document()
    document.add_heading("Python Scripts Documentation", level=0)

    current_folder = None

    for dirpath, fname in files:
        full_path = os.path.join(dirpath, fname)
        print(f"Running: {full_path}")

        # New folder heading whenever we move to a different folder
        if dirpath != current_folder:
            document.add_heading(f"Folder: {dirpath}", level=1)
            current_folder = dirpath

        document.add_heading(f"File: {fname}", level=2)

        # Source code
        document.add_paragraph("Code:", style="Intense Quote")
        try:
            with open(full_path, "r", encoding="utf-8", errors="replace") as f:
                source_code = f.read()
        except Exception as e:
            source_code = f"Could not read file: {e}"
        add_code_block(document, source_code)

        # Run and capture output
        status, output_text = run_script(full_path)
        label = {"OK": "Output:", "ERROR": "Error Output:", "TIMEOUT": "Timed Out:"}[
            status
        ]
        document.add_paragraph(label, style="Intense Quote")
        add_code_block(document, output_text)

        document.add_page_break()

    save_path = os.path.join(save_folder, "Python Scripts Documentation.docx")
    document.save(save_path)
    print(f"\nDone. Document saved to: {save_path}")


if __name__ == "__main__":
    main()
