#!/usr/bin/env python3
"""Fast regression tests for the dependency-free Office Open XML helper."""

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
HELPER = ROOT / "scripts" / "office-openxml.py"


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"""


def root_rels(target, rel_type):
    return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="%s" Target="%s"/>
</Relationships>""" % (rel_type, target)


def write_package(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES)
        for name, value in parts.items():
            package.writestr(name, value)


def run_helper(path, kind, action="inspect", **options):
    request = {"action": action, "inputPath": str(path), "kind": kind, **options}
    result = subprocess.run(
        [sys.executable, str(HELPER)],
        input=json.dumps(request),
        text=True,
        capture_output=True,
        check=False,
        env={**os.environ, "PYTHONPYCACHEPREFIX": str(path.parent / "pycache")},
    )
    parsed = json.loads(result.stdout or "{}")
    if result.returncode != 0 or not parsed.get("ok"):
        raise AssertionError(parsed.get("error") or result.stderr)
    return parsed["value"]


def emit_fixtures(root):
    root.mkdir(parents=True, exist_ok=True)
    docx = root / "sample.docx"
    write_package(docx, {
        "_rels/.rels": root_rels("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "word/document.xml": """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello Word</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="customer"/><w:alias w:val="Customer"/><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Acme</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>""",
    })
    xlsx = root / "sample.xlsx"
    write_package(xlsx, {
        "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Total">Data!$A$1</definedName></definedNames></workbook>""",
        "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
        "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>SUM(1,2)</f><v>3</v></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row></sheetData></worksheet>""",
    })
    pptx = root / "sample.pptx"
    write_package(pptx, {
        "_rels/.rels": root_rels("ppt/presentation.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "ppt/presentation.xml": """<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>""",
        "ppt/_rels/presentation.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""",
        "ppt/slides/slide1.xml": """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello PowerPoint</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>""",
    })
    return {"word": docx, "excel": xlsx, "powerpoint": pptx}


def main():
    checks = {}
    with tempfile.TemporaryDirectory(prefix="onedrive-office-test-") as directory:
        root = Path(directory)

        docx = root / "sample.docx"
        write_package(docx, {
            "_rels/.rels": root_rels("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "word/document.xml": """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello Word</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="customer"/><w:alias w:val="Customer"/><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Acme</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>""",
        })
        word = run_helper(docx, "word")
        checks["word"] = word["paragraphs"][0]["text"] == "Hello Word" and word["tableCount"] == 1 and word["contentControlCount"] == 1
        edited_docx = root / "edited.docx"
        word_edit = run_helper(docx, "word", action="edit", outputPath=str(edited_docx), operations=[{"type": "replaceText", "find": "Hello Word", "replace": "Edited Word"}])
        edited_word = run_helper(edited_docx, "word")
        checks["wordEdit"] = word_edit["changeCount"] == 1 and edited_word["paragraphs"][0]["text"] == "Edited Word"
        rich_docx = root / "rich.docx"
        rich_word_edit = run_helper(docx, "word", action="edit", outputPath=str(rich_docx), operations=[
            {"type": "setParagraphText", "paragraphIndex": 0, "text": "Native Word"},
            {"type": "setParagraphStyle", "paragraphIndex": 0, "style": "Title"},
            {"type": "setTableCell", "tableIndex": 0, "rowIndex": 0, "columnIndex": 0, "text": "Table value"},
            {"type": "setContentControlText", "tag": "customer", "text": "Contoso"},
            {"type": "insertParagraph", "text": "Inserted paragraph", "style": "Normal"},
        ])
        rich_word = run_helper(rich_docx, "word")
        checks["wordStructuredEdits"] = rich_word_edit["changeCount"] == 5 and rich_word["paragraphs"][0]["text"] == "Native Word" and rich_word["paragraphs"][0]["style"] == "Title" and rich_word["tables"][0]["rows"][0][0] == "Table value" and rich_word["contentControls"][0]["text"] == "Contoso" and any(paragraph["text"] == "Inserted paragraph" for paragraph in rich_word["paragraphs"])
        expanding_docx = root / "expanding.docx"
        expanding = run_helper(docx, "word", action="edit", outputPath=str(expanding_docx), operations=[{"type": "replaceText", "find": "o", "replace": "oo", "all": True}])
        checks["wordExpandingReplacementBounded"] = expanding["changeCount"] == 2

        xlsx = root / "sample.xlsx"
        write_package(xlsx, {
            "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Total">Data!$A$1</definedName></definedNames></workbook>""",
            "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
            "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>SUM(1,2)</f><v>3</v></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row></sheetData></worksheet>""",
        })
        excel = run_helper(xlsx, "excel")
        cells = excel["sheets"][0]["cells"]
        checks["excel"] = excel["sheetCount"] == 1 and cells[0]["formula"] == "SUM(1,2)" and cells[1]["value"] == "Revenue"
        edited_xlsx = root / "edited.xlsx"
        excel_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(edited_xlsx), operations=[{"type": "setCell", "sheet": "Data", "address": "B2", "value": "Updated"}, {"type": "setFormula", "sheet": "Data", "address": "C2", "formula": "1+2"}])
        edited_excel = run_helper(edited_xlsx, "excel")
        edited_cells = {cell["address"]: cell for cell in edited_excel["sheets"][0]["cells"]}
        checks["excelEdit"] = excel_edit["changeCount"] == 2 and edited_cells["B2"]["value"] == "Updated" and edited_cells["C2"]["formula"] == "1+2"
        rich_xlsx = root / "rich.xlsx"
        rich_excel_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(rich_xlsx), operations=[
            {"type": "setRange", "sheet": "Data", "address": "A3:B4", "values": [[1, 2], [3, 4]]},
            {"type": "setStyle", "sheet": "Data", "address": "A3:B3", "styleIndex": 2},
            {"type": "setDefinedName", "name": "InputBlock", "formula": "Data!$A$3:$B$4"},
            {"type": "renameSheet", "sheet": "Data", "newName": "Results"},
        ])
        rich_excel = run_helper(rich_xlsx, "excel")
        rich_cells = {cell["address"]: cell for cell in rich_excel["sheets"][0]["cells"]}
        checks["excelRangeAndMetadataEdits"] = rich_excel_edit["changeCount"] == 8 and rich_excel["sheets"][0]["name"] == "Results" and rich_cells["B4"]["value"] == 4 and rich_cells["A3"]["styleIndex"] == 2 and any(entry["name"] == "InputBlock" for entry in rich_excel["definedNames"])

        pptx = root / "sample.pptx"
        write_package(pptx, {
            "_rels/.rels": root_rels("ppt/presentation.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "ppt/presentation.xml": """<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>""",
            "ppt/_rels/presentation.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""",
            "ppt/slides/slide1.xml": """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello PowerPoint</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>""",
        })
        powerpoint = run_helper(pptx, "powerpoint")
        checks["powerpoint"] = powerpoint["slideCount"] == 1 and powerpoint["slides"][0]["shapes"][0]["text"] == "Hello PowerPoint"
        edited_pptx = root / "edited.pptx"
        ppt_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(edited_pptx), operations=[{"type": "replaceText", "slideIndex": 0, "shapeId": "2", "find": "Hello", "replace": "Edited"}])
        edited_powerpoint = run_helper(edited_pptx, "powerpoint")
        checks["powerpointEdit"] = ppt_edit["changeCount"] == 1 and edited_powerpoint["slides"][0]["shapes"][0]["text"] == "Edited PowerPoint"
        rich_pptx = root / "rich.pptx"
        rich_ppt_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(rich_pptx), operations=[
            {"type": "setShapeText", "slideIndex": 0, "shapeId": "2", "text": "Native PowerPoint"},
            {"type": "setShapeGeometry", "slideIndex": 0, "shapeId": "2", "x": 10, "y": 20, "width": 30, "height": 40},
        ])
        rich_powerpoint = run_helper(rich_pptx, "powerpoint")
        shape = rich_powerpoint["slides"][0]["shapes"][0]
        checks["powerpointStructuredEdits"] = rich_ppt_edit["changeCount"] == 2 and shape["text"] == "Native PowerPoint" and all(shape["geometry"][key] == value for key, value in {"x": 10, "y": 20, "width": 30, "height": 40}.items())
        slides_pptx = root / "slides.pptx"
        slides_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(slides_pptx), operations=[{"type": "duplicateSlide", "slideIndex": 0}, {"type": "moveSlide", "slideIndex": 1, "toIndex": 0}])
        slides_powerpoint = run_helper(slides_pptx, "powerpoint")
        checks["powerpointSlideLifecycle"] = slides_edit["changeCount"] == 2 and slides_powerpoint["slideCount"] == 2 and all(slide["shapes"][0]["text"] == "Hello PowerPoint" for slide in slides_powerpoint["slides"])

        bad = root / "bad.docx"
        write_package(bad, {
            "_rels/.rels": root_rels("word/missing.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "word/document.xml": "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>",
        })
        bad_result = subprocess.run(
            [sys.executable, str(HELPER)],
            input=json.dumps({"action": "validate", "inputPath": str(bad), "kind": "word"}),
            text=True,
            capture_output=True,
            check=False,
        )
        checks["brokenRelationshipRejected"] = bad_result.returncode != 0 and "broken internal relationships" in bad_result.stdout
        wrong_kind = subprocess.run(
            [sys.executable, str(HELPER)],
            input=json.dumps({"action": "edit", "inputPath": str(docx), "outputPath": str(root / "wrong.pptx"), "kind": "powerpoint", "operations": [{"type": "replaceText", "find": "Hello", "replace": "Wrong"}]}),
            text=True,
            capture_output=True,
            check=False,
        )
        checks["wrongKindEditRejected"] = wrong_kind.returncode != 0 and "Expected a powerpoint package" in wrong_kind.stdout
        signed = root / "signed.docx"
        write_package(signed, {
            "_rels/.rels": root_rels("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "word/document.xml": "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Signed</w:t></w:r></w:p></w:body></w:document>",
            "_xmlsignatures/sig1.xml": "<Signature/>",
        })
        signed_result = subprocess.run(
            [sys.executable, str(HELPER)], input=json.dumps({"action": "edit", "inputPath": str(signed), "outputPath": str(root / "signed-edit.docx"), "kind": "word", "allowSignedPackage": True, "operations": [{"type": "replaceText", "find": "Signed", "replace": "Changed"}]}), text=True, capture_output=True, check=False
        )
        checks["signedEditAlwaysRejected"] = signed_result.returncode != 0 and "Digitally signed" in signed_result.stdout
        macro = root / "macro.docx"
        write_package(macro, {
            "_rels/.rels": root_rels("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "word/document.xml": "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Macro</w:t></w:r></w:p></w:body></w:document>",
            "word/vbaProject.bin": b"mock-vba-never-executed",
        })
        macro_result = subprocess.run(
            [sys.executable, str(HELPER)], input=json.dumps({"action": "edit", "inputPath": str(macro), "outputPath": str(root / "macro-edit.docx"), "kind": "word", "operations": [{"type": "replaceText", "find": "Macro", "replace": "Changed"}]}), text=True, capture_output=True, check=False
        )
        checks["macroEditRequiresOptIn"] = macro_result.returncode != 0 and "allowMacros=true" in macro_result.stdout

    ok = all(checks.values())
    print(json.dumps({"ok": ok, "checks": checks}, indent=2))
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    fixture_arg = next((arg for arg in sys.argv[1:] if arg.startswith("--emit-fixtures=")), None)
    if fixture_arg:
        fixtures = emit_fixtures(Path(fixture_arg.split("=", 1)[1]).resolve())
        print(json.dumps({key: str(value) for key, value in fixtures.items()}))
    else:
        main()
