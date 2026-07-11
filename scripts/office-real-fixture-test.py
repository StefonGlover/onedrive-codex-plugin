#!/usr/bin/env python3
"""Interoperability QA using real Office packages and LibreOffice rendering."""

import json
import os
import subprocess
import sys
import tempfile
from contextlib import nullcontext
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation
from pptx.util import Inches


ROOT = Path(__file__).resolve().parents[1]
HELPER = ROOT / "scripts" / "office-openxml.py"
SOFFICE = Path(os.environ.get("SOFFICE", "/Users/stefonglover/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice"))


def edit(source, destination, kind, operations):
    result = subprocess.run(
        ["/usr/bin/python3", str(HELPER)],
        input=json.dumps({"action": "edit", "inputPath": str(source), "outputPath": str(destination), "kind": kind, "operations": operations}),
        text=True,
        capture_output=True,
        check=False,
        env={**os.environ, "PYTHONPYCACHEPREFIX": str(destination.parent / "pycache")},
    )
    payload = json.loads(result.stdout or "{}")
    if result.returncode or not payload.get("ok"):
        raise RuntimeError(payload.get("error") or result.stderr)
    return payload["value"]


def main():
    checks = {}
    kept_root = os.environ.get("OFFICE_REAL_FIXTURE_DIR")
    context = nullcontext(kept_root) if kept_root else tempfile.TemporaryDirectory(prefix="onedrive-office-real-")
    with context as directory:
        root = Path(directory)
        root.mkdir(parents=True, exist_ok=True)
        output = root / "rendered"
        output.mkdir()

        document = Document()
        document.add_heading("Real Word Fixture", level=1)
        document.add_paragraph("Hello from Word")
        source_docx = root / "real.docx"
        document.save(source_docx)
        edited_docx = root / "edited-word.docx"
        word = edit(source_docx, edited_docx, "word", [{"type": "replaceText", "find": "Hello from Word", "replace": "Edited in OneDrive"}, {"type": "insertParagraph", "text": "Inserted safely"}])

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data"
        sheet["A1"] = "Revenue"
        sheet["B1"] = 10
        source_xlsx = root / "real.xlsx"
        workbook.save(source_xlsx)
        edited_xlsx = root / "edited-excel.xlsx"
        excel = edit(source_xlsx, edited_xlsx, "excel", [{"type": "setRange", "sheet": "Data", "address": "A2:B3", "values": [["North", 20], ["South", 30]]}])

        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        box.text = "Hello from PowerPoint"
        source_pptx = root / "real.pptx"
        presentation.save(source_pptx)
        edited_pptx = root / "edited-powerpoint.pptx"
        powerpoint = edit(source_pptx, edited_pptx, "powerpoint", [{"type": "setShapeText", "slideIndex": 0, "shapeId": str(box.shape_id), "text": "Edited in OneDrive"}, {"type": "duplicateSlide", "slideIndex": 0}])

        Document(edited_docx)
        load_workbook(edited_xlsx)
        Presentation(edited_pptx)
        rendered = {}
        for path in (edited_docx, edited_xlsx, edited_pptx):
            profile = root / ("profile-" + path.stem)
            conversion = subprocess.run([str(SOFFICE), "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf", "--outdir", str(output), str(path)], cwd=root, capture_output=True, text=True, check=False)
            pdf = output / (path.stem + ".pdf")
            rendered[path.suffix] = {"exitCode": conversion.returncode, "bytes": pdf.stat().st_size if pdf.exists() else 0, "stderr": conversion.stderr[-500:]}

        checks["wordRealPackage"] = word["changeCount"] == 2 and rendered[".docx"]["bytes"] > 0
        checks["excelRealPackage"] = excel["changeCount"] == 4 and rendered[".xlsx"]["bytes"] > 0
        checks["powerpointRealPackage"] = powerpoint["changeCount"] == 2 and rendered[".pptx"]["bytes"] > 0
        print(json.dumps({"ok": all(checks.values()), "checks": checks, "rendered": rendered}, indent=2))
        raise SystemExit(0 if all(checks.values()) else 1)


if __name__ == "__main__":
    main()
