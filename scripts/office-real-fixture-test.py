#!/usr/bin/env python3
"""Interoperability QA using real Office packages and LibreOffice rendering."""

import json
import os
import subprocess
import sys
import tempfile
import base64
import shutil
from io import BytesIO
from contextlib import nullcontext
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from pptx import Presentation
from pptx.util import Inches
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
HELPER = ROOT / "scripts" / "office-openxml.py"
SOFFICE = Path(os.environ.get("SOFFICE") or shutil.which("soffice") or "soffice")


def edit(source, destination, kind, operations):
    result = subprocess.run(
        [sys.executable, str(HELPER)],
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
        sheet.append(["Region", "Revenue"])
        sheet.append(["North", 20])
        sheet.append(["South", 30])
        table = Table(displayName="RevenueTable", ref="A1:B3")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)
        source_xlsx = root / "real.xlsx"
        workbook.save(source_xlsx)
        edited_xlsx = root / "edited-excel.xlsx"
        excel = edit(source_xlsx, edited_xlsx, "excel", [
            {"type": "addTableRow", "table": "RevenueTable", "values": [["West", 40]]},
            {"type": "setTableTotals", "table": "RevenueTable", "enabled": True, "columns": [{"column": "Region", "label": "Total"}, {"column": "Revenue", "function": "sum"}]},
            {"type": "createChart", "sheet": "Data", "sourceData": "A1:B4", "chartType": "ColumnClustered", "name": "Revenue chart", "titleText": "Revenue by region", "left": 20, "top": 30, "width": 420, "height": 240},
            {"type": "updateChart", "sheet": "Data", "chart": "Revenue chart", "chartType": "Line", "sourceData": "A1:B4", "name": "Revenue trend", "titleText": "Revenue trend", "left": 40, "top": 50, "width": 400, "height": 220},
        ])

        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        box.text = "Hello from PowerPoint"
        disposable = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(0.5))
        disposable.text = "Delete me"
        source_png_buffer = BytesIO()
        Image.new("RGB", (8, 8), (37, 99, 235)).save(source_png_buffer, format="PNG")
        source_png = source_png_buffer.getvalue()
        picture = slide.shapes.add_picture(BytesIO(source_png), Inches(1), Inches(3), Inches(1), Inches(1))
        source_pptx = root / "real.pptx"
        presentation.save(source_pptx)
        edited_pptx = root / "edited-powerpoint.pptx"
        powerpoint = edit(source_pptx, edited_pptx, "powerpoint", [
            {"type": "setShapeText", "slideIndex": 0, "shapeId": str(box.shape_id), "text": "Edited in OneDrive"},
            {"type": "setTextStyle", "slideIndex": 0, "shapeId": str(box.shape_id), "bold": True, "fontSize": 24, "color": "2563EB"},
            {"type": "addTextBox", "slideIndex": 0, "text": "Native text box", "x": 914400, "y": 4114800, "width": 2743200, "height": 457200},
            {"type": "deleteShape", "slideIndex": 0, "shapeId": str(disposable.shape_id)},
            {"type": "replaceImage", "slideIndex": 0, "shapeId": str(picture.shape_id), "base64": base64.b64encode(source_png).decode("ascii"), "contentType": "image/png"},
            {"type": "duplicateSlide", "slideIndex": 0},
        ])

        Document(edited_docx)
        reopened_excel = load_workbook(edited_xlsx, data_only=False)
        Presentation(edited_pptx)
        rendered = {}
        for path in (edited_docx, edited_xlsx, edited_pptx):
            profile = root / ("profile-" + path.stem)
            conversion = subprocess.run([str(SOFFICE), "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf", "--outdir", str(output), str(path)], cwd=root, capture_output=True, text=True, check=False)
            pdf = output / (path.stem + ".pdf")
            diagnostic = (conversion.stdout + "\n" + conversion.stderr).lower()
            rendered[path.suffix] = {"exitCode": conversion.returncode, "bytes": pdf.stat().st_size if pdf.exists() else 0, "cleanOpen": conversion.returncode == 0 and not any(term in diagnostic for term in ("repair", "corrupt", "fatal error")), "stderr": conversion.stderr[-500:]}

        checks["wordRealPackage"] = word["changeCount"] == 2 and rendered[".docx"]["bytes"] > 0 and rendered[".docx"]["cleanOpen"]
        reopened_sheet = reopened_excel["Data"]
        checks["excelRealPackage"] = (
            excel["changeCount"] == 4
            and reopened_sheet.tables["RevenueTable"].ref == "A1:B5"
            and len(reopened_sheet._charts) == 1
            and reopened_sheet["B5"].value == "=SUBTOTAL(109,[Revenue])"
            and rendered[".xlsx"]["bytes"] > 0
            and rendered[".xlsx"]["cleanOpen"]
        )
        checks["powerpointRealPackage"] = powerpoint["changeCount"] == 6 and rendered[".pptx"]["bytes"] > 0 and rendered[".pptx"]["cleanOpen"]
        print(json.dumps({"ok": all(checks.values()), "checks": checks, "rendered": rendered}, indent=2))
        raise SystemExit(0 if all(checks.values()) else 1)


if __name__ == "__main__":
    main()
