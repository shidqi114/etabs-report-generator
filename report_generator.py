"""
report_generator.py
===================
Word document generation layer for the ETABS Report Generator.

Workflow
────────
  1. DocxTemplate (docxtpl / Jinja2) renders all placeholders and table
     loops into the tagged .docx template.
  2. The document is saved to disk as a new _Generated.docx file.
  3. update_word_toc() silently opens the saved document via the Word COM
     API to refresh all Table-of-Contents field codes and page numbers,
     then saves and closes the file.  On non-Windows systems the step is
     skipped gracefully with a DEV MODE notice.

Spec reference: "Report generation tagging.md"
"""

from __future__ import annotations

import os
import platform
from typing import Any

from docxtpl import DocxTemplate


# ---------------------------------------------------------------------------
# TOC updater (Step 3 from spec: Report generation tagging.md)
# ---------------------------------------------------------------------------

def update_word_toc(file_path: str) -> None:
    """
    Silently opens Microsoft Word via COM to update all Table of Contents
    field codes and page numbers.

    The document must contain a *native* Word automatic TOC (inserted via
    References → Table of Contents → Automatic Table 1/2) for this to work.
    Plain-text lists that look like a TOC will be ignored.

    On macOS / Linux this function exits immediately with a DEV MODE notice;
    no Word COM dependency is required for cross-platform development.

    Parameters
    ----------
    file_path : str
        Absolute path to the generated .docx file.
    """
    if platform.system() != "Windows":
        print("DEV MODE: Skipping TOC update — requires Windows COM (Word).")
        return

    # win32com ships with pywin32 (pip install pywin32)
    import win32com.client  # type: ignore[import]

    abs_path = os.path.abspath(file_path)
    word = None
    try:
        # DispatchEx creates a fully isolated, hidden Word instance
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(abs_path)

        if doc.TablesOfContents.Count > 0:
            for toc in doc.TablesOfContents:
                toc.Update()
            print(f"[report_generator] Table of Contents updated in: {abs_path}")
        else:
            print("[report_generator] No automatic Table of Contents found — skipping TOC update.")

        doc.Close(SaveChanges=True)

    except Exception as e:
        # Non-fatal: report is still usable; user can manually refresh the TOC
        print(f"[report_generator] Warning: Could not update TOC automatically. Error: {e}")

    finally:
        # Guarantee the hidden Word process is terminated even on error
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Main report generation function
# ---------------------------------------------------------------------------

def generate_report(
    context_data: dict[str, Any],
    template_path: str,
    output_path: str,
) -> str:
    """
    Renders the tagged Word template with ETABS data and saves the output.

    Parameters
    ----------
    context_data : dict
        Full context dictionary produced by ``data_pipeline.build_report_context()``.
        Contains all Jinja2 variables referenced in the .docx template tags.
    template_path : str
        Path to the tagged Word template (e.g. Report_Template_Tagged.docx).
    output_path : str
        Destination path for the generated document (e.g. *_Generated.docx).

    Returns
    -------
    str
        Absolute path to the generated document.
    """
    print("[report_generator] Initialising Word template engine...")

    # 1. Load the tagged .docx template
    doc = DocxTemplate(template_path)

    # 2. Inject all context variables (Jinja2 renders placeholders + table loops)
    print("[report_generator] Rendering data into template...")
    doc.render(context_data)

    # 3. Save the XML-modified document to disk
    abs_output = os.path.abspath(output_path)
    doc.save(abs_output)
    print(f"[report_generator] Document saved → {abs_output}")

    # 4. Update TOC field codes and page numbers via Word COM (Windows only)
    print("[report_generator] Finalising page numbers and Table of Contents...")
    update_word_toc(abs_output)

    print("[report_generator] ✅ Report Generation Complete!")
    return abs_output


if __name__ == "__main__":
    print("Report Generator module ready.")
