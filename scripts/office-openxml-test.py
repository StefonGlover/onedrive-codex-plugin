#!/usr/bin/env python3
"""Fast regression tests for the dependency-free Office Open XML helper."""

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
HELPER = ROOT / "scripts" / "office-openxml.py"

OPENXML_TOOL_NAMES = {
    "word": "onedrive_word_batch_update",
    "excel": "onedrive_excel_batch_update",
    "powerpoint": "onedrive_powerpoint_batch_update",
}
EXPECTED_OPENXML_OPERATION_COUNTS = {"word": 21, "excel": 33, "powerpoint": 25}
COVERED_OPENXML_OPERATIONS = {kind: set() for kind in OPENXML_TOOL_NAMES}
RICH_REAL_FIXTURE_OPERATIONS = {
    "word": {"insertImage", "replaceImage", "createContentControl", "deleteContentControl", "createBookmark", "deleteBookmark", "insertTableRow", "deleteTableRow", "insertTableColumn", "deleteTableColumn", "setHeaderFooterText", "setSectionProperties"},
    "excel": {"addWorksheet", "deleteWorksheet", "addTable", "deleteTable", "mergeRange", "unmergeRange", "sortRange", "setAutoFilter", "setHyperlink", "addNote", "deleteNote", "insertImage", "formatChart", "setSheetProtection", "refreshPivot"},
    "powerpoint": {"addSlide", "addImage", "cropImage", "addTable", "insertTableRow", "deleteTableRow", "insertTableColumn", "deleteTableColumn", "setShapeAltText", "setZOrder", "groupShapes", "ungroupShape", "applySlideLayout"},
}


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"""

WORD_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

EXCEL_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"""


def root_rels(target, rel_type):
    return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="%s" Target="%s"/>
</Relationships>""" % (rel_type, target)


def write_package(path, parts, content_types=CONTENT_TYPES):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", content_types)
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
    if action == "edit":
        COVERED_OPENXML_OPERATIONS[kind].update(operation.get("type") for operation in options.get("operations", []))
    return parsed["value"]


def production_openxml_operations():
    """Read the advertised operation contracts from the production MCP schemas."""
    requests = "\n".join([
        json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "office-openxml-test", "version": "1"}}}),
        json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}}),
        "",
    ])
    result = subprocess.run(
        ["node", str(ROOT / "mcp" / "server.mjs")],
        cwd=ROOT,
        input=requests,
        text=True,
        capture_output=True,
        check=False,
        timeout=10,
        env={**os.environ, "ONEDRIVE_TEST_ACCESS_TOKEN": "office-openxml-schema-check"},
    )
    if result.returncode != 0:
        raise AssertionError(result.stderr or "MCP schema inspection failed")
    messages = [json.loads(line) for line in result.stdout.splitlines() if line.strip()]
    tools = next(message["result"]["tools"] for message in messages if message.get("id") == 2)
    tools_by_name = {tool["name"]: tool for tool in tools}
    contract = {}
    for kind, tool_name in OPENXML_TOOL_NAMES.items():
        operation_variants = tools_by_name[tool_name]["inputSchema"]["properties"]["operations"]["items"]["anyOf"]
        contract[kind] = {variant["properties"]["type"]["const"] for variant in operation_variants}
    return contract


def emit_fixtures(root):
    root.mkdir(parents=True, exist_ok=True)
    docx = root / "sample.docx"
    write_package(docx, {
        "_rels/.rels": root_rels("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "word/document.xml": """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello Word</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="customer"/><w:alias w:val="Customer"/><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Acme</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>""",
    }, WORD_CONTENT_TYPES)
    xlsx = root / "sample.xlsx"
    write_package(xlsx, {
        "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Total">Data!$A$1</definedName></definedNames></workbook>""",
        "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>""",
        "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><f>SUM(1,2)</f><v>3</v></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts><drawing r:id="rId2"/></worksheet>""",
        "xl/worksheets/_rels/sheet1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>""",
        "xl/tables/table1.xml": """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="RevenueTable" displayName="RevenueTable" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="Metric"/><tableColumn id="2" name="Revenue"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>""",
        "xl/drawings/drawing1.xml": """<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><c:chart r:id="rId1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>""",
        "xl/drawings/_rels/drawing1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>""",
        "xl/charts/chart1.xml": """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue Chart</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>""",
        "xl/styles.xml": """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="0"/><fonts count="1"><font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>""",
        "xl/calcChain.xml": """<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>""",
    }, EXCEL_CONTENT_TYPES)
    pptx = root / "sample.pptx"
    ppt_content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>"""
    write_package(pptx, {
        "_rels/.rels": root_rels("ppt/presentation.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "ppt/presentation.xml": """<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>""",
        "ppt/_rels/presentation.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""",
        "ppt/slides/slide1.xml": """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello PowerPoint</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="300"/></a:tblGrid><a:tr h="200"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old table value</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>""",
        "ppt/slides/_rels/slide1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.gif"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>""",
        "ppt/notesSlides/notesSlide1.xml": """<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>""",
        "ppt/media/image1.gif": b"GIF89a",
    }, ppt_content_types)
    return {"word": docx, "excel": xlsx, "powerpoint": pptx}


def main():
    checks = {}
    with tempfile.TemporaryDirectory(prefix="onedrive-office-test-") as directory:
        root = Path(directory)

        fixtures = emit_fixtures(root)
        docx = fixtures["word"]
        word = run_helper(docx, "word", searchText="Hello")
        checks["word"] = word["paragraphs"][0]["text"] == "Hello Word" and word["tableCount"] == 1 and word["contentControlCount"] == 1 and word["search"]["matchCount"] == 1
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
            {"type": "addHyperlink", "paragraphIndex": 0, "text": " OpenAI", "url": "https://openai.com/docs"},
            {"type": "addComment", "paragraphIndex": 0, "text": "Review this paragraph", "author": "Codex", "initials": "CX"},
            {"type": "insertTable", "afterParagraphIndex": 0, "rows": [["Name", "Value"], ["Alpha", "1"]], "style": "TableGrid"},
        ])
        rich_word = run_helper(rich_docx, "word")
        checks["wordStructuredEdits"] = rich_word_edit["changeCount"] == 8 and rich_word["paragraphs"][0]["text"] == "Native Word OpenAI" and rich_word["paragraphs"][0]["style"] == "Title" and rich_word["tables"][0]["rows"][0] == ["Name", "Value"] and rich_word["tables"][1]["rows"][0][0] == "Table value" and rich_word["contentControls"][0]["text"] == "Contoso" and rich_word["comments"][0]["text"] == "Review this paragraph" and any(paragraph["text"] == "Inserted paragraph" for paragraph in rich_word["paragraphs"])
        with zipfile.ZipFile(rich_docx, "r") as package:
            document_xml = package.read("word/document.xml").decode("utf-8")
            relationships_xml = package.read("word/_rels/document.xml.rels").decode("utf-8")
            content_types_xml = package.read("[Content_Types].xml").decode("utf-8")
            checks["wordHyperlinkRelationship"] = "https://openai.com/docs" in relationships_xml and 'TargetMode="External"' in relationships_xml and "hyperlink" in document_xml
            checks["wordCommentPackageParts"] = "word/comments.xml" in package.namelist() and "comments.xml" in relationships_xml and "/word/comments.xml" in content_types_xml and "commentRangeStart" in document_xml and "commentReference" in document_xml
        expanding_docx = root / "expanding.docx"
        expanding = run_helper(docx, "word", action="edit", outputPath=str(expanding_docx), operations=[{"type": "replaceText", "find": "o", "replace": "oo", "all": True}])
        checks["wordExpandingReplacementBounded"] = expanding["changeCount"] == 2
        tracked_docx = root / "tracked.docx"
        write_package(tracked_docx, {
            "_rels/.rels": root_rels("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "word/document.xml": """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins w:id="1"><w:r><w:t>Tracked text</w:t></w:r></w:ins></w:p></w:body></w:document>""",
        })
        tracked_result = subprocess.run(
            [sys.executable, str(HELPER)],
            input=json.dumps({"action": "edit", "inputPath": str(tracked_docx), "outputPath": str(root / "tracked-edited.docx"), "kind": "word", "operations": [{"type": "insertParagraph", "text": "Unsafe"}]}),
            text=True,
            capture_output=True,
            check=False,
        )
        checks["wordTrackedChangesRefused"] = tracked_result.returncode != 0 and "tracked changes are refused" in tracked_result.stdout

        xlsx = fixtures["excel"]
        excel = run_helper(xlsx, "excel", searchText="Revenue")
        cells = excel["sheets"][0]["cells"]
        checks["excel"] = excel["sheetCount"] == 1 and cells[0]["value"] == "Revenue" and excel["search"]["matchCount"] == 1 and excel["tableCount"] == 1 and excel["sheets"][0]["tables"][0]["displayName"] == "RevenueTable" and excel["chartCount"] == 1 and excel["sheets"][0]["charts"][0]["title"] == "Revenue Chart" and excel["pivotCount"] == 0
        selected_excel = run_helper(xlsx, "excel", sheetNames=["Data"], address="A1:A1")
        checks["excelSelectors"] = selected_excel["cellCount"] == 1 and selected_excel["sheets"][0]["cells"][0]["formula"] == "SUM(1,2)"
        edited_xlsx = root / "edited.xlsx"
        excel_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(edited_xlsx), operations=[{"type": "setCell", "sheet": "Data", "address": "B2", "value": "Updated"}, {"type": "setFormula", "sheet": "Data", "address": "C2", "formula": "B2+1"}])
        edited_excel = run_helper(edited_xlsx, "excel", includeFormulaDependencies=True)
        edited_cells = {cell["address"]: cell for cell in edited_excel["sheets"][0]["cells"]}
        checks["excelEdit"] = excel_edit["changeCount"] == 2 and edited_cells["B2"]["value"] == "Updated" and edited_cells["C2"]["formula"] == "B2+1" and edited_excel["formulaDependencyCount"] == 1 and edited_excel["formulaDependencies"][0]["to"]["address"] == "B2"
        rich_xlsx = root / "rich.xlsx"
        rich_excel_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(rich_xlsx), operations=[
            {"type": "setRange", "sheet": "Data", "address": "A3:B4", "values": [[1, 2], [3, 4]]},
            {"type": "setStyle", "sheet": "Data", "address": "A3:B3", "styleIndex": 2},
            {"type": "setNumberFormat", "sheet": "Data", "address": "A1", "formatCode": "$#,##0.00"},
            {"type": "addConditionalFormat", "sheet": "Data", "address": "A3:B4", "ruleType": "cellIs", "operator": "greaterThan", "formula": "2", "fillColor": "FFF2CC"},
            {"type": "setDataValidation", "sheet": "Data", "address": "A3:A4", "validationType": "whole", "operator": "between", "formula1": "1", "formula2": "10"},
            {"type": "freezePanes", "sheet": "Data", "rows": 1, "columns": 1},
            {"type": "setColumnWidth", "sheet": "Data", "address": "A1:B1", "width": 18},
            {"type": "setDefinedName", "name": "InputBlock", "formula": "Data!$A$3:$B$4"},
            {"type": "renameSheet", "sheet": "Data", "newName": "Results"},
            {"type": "recalculate"},
        ])
        rich_excel = run_helper(rich_xlsx, "excel")
        rich_cells = {cell["address"]: cell for cell in rich_excel["sheets"][0]["cells"]}
        with zipfile.ZipFile(rich_xlsx, "r") as edited_package:
            workbook_xml = edited_package.read("xl/workbook.xml").decode("utf-8")
            styles_xml = edited_package.read("xl/styles.xml").decode("utf-8")
            relations_xml = edited_package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_xml = edited_package.read("xl/worksheets/sheet1.xml").decode("utf-8")
            recalculation_safe = "fullCalcOnLoad=\"1\"" in workbook_xml and "forceFullCalc=\"1\"" in workbook_xml and "xl/calcChain.xml" not in edited_package.namelist() and "calcChain" not in relations_xml
        checks["excelRangeAndMetadataEdits"] = rich_excel_edit["changeCount"] == 14 and rich_excel["sheets"][0]["name"] == "Results" and rich_cells["B4"]["value"] == 4 and rich_cells["A3"]["styleIndex"] == 2 and rich_cells["A1"]["styleIndex"] == 3 and "$#,##0.00" in styles_xml and "conditionalFormatting" in sheet_xml and "dataValidation" in sheet_xml and "state=\"frozen\"" in sheet_xml and "width=\"18.0\"" in sheet_xml and recalculation_safe and any(entry["name"] == "InputBlock" for entry in rich_excel["definedNames"])

        styled_xlsx = root / "styled-for-clear.xlsx"
        run_helper(xlsx, "excel", action="edit", outputPath=str(styled_xlsx), operations=[
            {"type": "setStyle", "sheet": "Data", "address": "B2", "styleIndex": 2},
        ])
        clear_contents_xlsx = root / "clear-contents.xlsx"
        clear_contents_edit = run_helper(styled_xlsx, "excel", action="edit", outputPath=str(clear_contents_xlsx), operations=[
            {"type": "clearRange", "sheet": "Data", "address": "B2", "contents": True, "format": False},
        ])
        clear_contents_cells = {cell["address"]: cell for cell in run_helper(clear_contents_xlsx, "excel")["sheets"][0]["cells"]}
        clear_format_xlsx = root / "clear-format.xlsx"
        clear_format_edit = run_helper(styled_xlsx, "excel", action="edit", outputPath=str(clear_format_xlsx), operations=[
            {"type": "clearRange", "sheet": "Data", "address": "B2", "contents": False, "format": True},
        ])
        clear_format_cells = {cell["address"]: cell for cell in run_helper(clear_format_xlsx, "excel")["sheets"][0]["cells"]}
        checks["excelClearRangeModes"] = (
            clear_contents_edit["changeCount"] == 1
            and clear_contents_cells["B2"]["value"] is None
            and clear_contents_cells["B2"]["styleIndex"] == 2
            and clear_format_edit["changeCount"] == 1
            and clear_format_cells["B2"]["value"] == 10
            and clear_format_cells["B2"]["styleIndex"] == 0
        )
        no_op_clear = subprocess.run(
            [sys.executable, str(HELPER)],
            input=json.dumps({"action": "edit", "inputPath": str(xlsx), "outputPath": str(root / "no-op-clear.xlsx"), "kind": "excel", "operations": [{"type": "clearRange", "sheet": "Data", "address": "B2", "contents": False, "format": False}]}),
            text=True,
            capture_output=True,
            check=False,
        )
        checks["excelClearRangeNoOpRejected"] = no_op_clear.returncode != 0 and "must clear contents, format, or both" in no_op_clear.stdout

        created_chart_xlsx = root / "chart-created.xlsx"
        create_chart_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(created_chart_xlsx), operations=[
            {"type": "createChart", "sheet": "Data", "chartType": "ColumnClustered", "sourceData": "A1:C3", "seriesBy": "Columns", "name": "Consumer Chart", "titleText": "Consumer Revenue", "left": 12, "top": 18, "width": 360, "height": 220},
        ])
        created_chart_result = run_helper(created_chart_xlsx, "excel")
        created_chart = next(chart for chart in created_chart_result["sheets"][0]["charts"] if chart.get("name") == "Consumer Chart")
        column_series_formulas = [entry["formulas"] for entry in created_chart["series"]]
        checks["excelCreateChartOptions"] = (
            create_chart_edit["changeCount"] == 1
            and created_chart["type"] == "barChart"
            and created_chart["title"] == "Consumer Revenue"
            and created_chart["seriesCount"] == 2
            and all(round(created_chart["geometry"][key]) == value for key, value in {"left": 12, "top": 18, "width": 360, "height": 220}.items())
        )
        checks["excelChartSeriesByColumns"] = column_series_formulas == [
            ["'Data'!$B$1", "'Data'!$A$2:$A$3", "'Data'!$B$2:$B$3"],
            ["'Data'!$C$1", "'Data'!$A$2:$A$3", "'Data'!$C$2:$C$3"],
        ]
        updated_chart_xlsx = root / "chart-updated.xlsx"
        update_chart_edit = run_helper(created_chart_xlsx, "excel", action="edit", outputPath=str(updated_chart_xlsx), operations=[
            {"type": "updateChart", "sheet": "Data", "chart": "Consumer Chart", "chartType": "Line", "sourceData": "A1:C3", "seriesBy": "Rows", "name": "Consumer Trend", "titleText": "Revenue Trend", "left": 20, "top": 24, "width": 400, "height": 240},
        ])
        updated_chart_result = run_helper(updated_chart_xlsx, "excel")
        updated_chart = next(chart for chart in updated_chart_result["sheets"][0]["charts"] if chart.get("name") == "Consumer Trend")
        row_series_formulas = [entry["formulas"] for entry in updated_chart["series"]]
        checks["excelUpdateChartOptions"] = (
            update_chart_edit["changeCount"] == 1
            and updated_chart["type"] == "lineChart"
            and updated_chart["title"] == "Revenue Trend"
            and updated_chart["seriesCount"] == 2
            and all(round(updated_chart["geometry"][key]) == value for key, value in {"left": 20, "top": 24, "width": 400, "height": 240}.items())
        )
        checks["excelChartSeriesByRows"] = row_series_formulas == [
            ["'Data'!$A$2", "'Data'!$B$1:$C$1", "'Data'!$B$2:$C$2"],
            ["'Data'!$A$3", "'Data'!$B$1:$C$1", "'Data'!$B$3:$C$3"],
        ]
        chart_type_only_xlsx = root / "chart-type-only.xlsx"
        chart_type_only_edit = run_helper(updated_chart_xlsx, "excel", action="edit", outputPath=str(chart_type_only_xlsx), operations=[
            {"type": "updateChart", "sheet": "Data", "chart": "Consumer Trend", "chartType": "Pie"},
        ])
        chart_type_only_result = run_helper(chart_type_only_xlsx, "excel")
        chart_type_only = next(chart for chart in chart_type_only_result["sheets"][0]["charts"] if chart.get("name") == "Consumer Trend")
        checks["excelChartTypeOnlyPreservesSeries"] = (
            chart_type_only_edit["changeCount"] == 1
            and chart_type_only["type"] == "pieChart"
            and [entry["formulas"] for entry in chart_type_only["series"]] == row_series_formulas
        )
        series_by_only_xlsx = root / "chart-series-by-only.xlsx"
        series_by_only_edit = run_helper(chart_type_only_xlsx, "excel", action="edit", outputPath=str(series_by_only_xlsx), operations=[
            {"type": "updateChart", "sheet": "Data", "chart": "Consumer Trend", "seriesBy": "Columns"},
        ])
        series_by_only_result = run_helper(series_by_only_xlsx, "excel")
        series_by_only = next(chart for chart in series_by_only_result["sheets"][0]["charts"] if chart.get("name") == "Consumer Trend")
        checks["excelChartSeriesByOnlyInfersSource"] = (
            series_by_only_edit["changeCount"] == 1
            and series_by_only["type"] == "pieChart"
            and [entry["formulas"] for entry in series_by_only["series"]] == column_series_formulas
        )

        table_xlsx = root / "table-operations.xlsx"
        table_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(table_xlsx), operations=[
            {"type": "setTableTotals", "table": "RevenueTable", "enabled": True, "columns": [
                {"column": "Metric", "label": "Total"},
                {"column": "Revenue", "function": "sum"},
            ]},
            {"type": "addTableRow", "table": "RevenueTable", "index": 0, "values": [["Q2", 20], ["Q3", 30]]},
        ])
        table_result = run_helper(table_xlsx, "excel")
        table_cells = {cell["address"]: cell for cell in table_result["sheets"][0]["cells"]}
        table_metadata = table_result["sheets"][0]["tables"][0]
        checks["excelTableRowsAndTotals"] = (
            table_edit["changeCount"] == 2
            and table_metadata["ref"] == "A1:B5"
            and table_metadata["totalsRowCount"] == 1
            and table_metadata["columns"][0]["totalsRowLabel"] == "Total"
            and table_metadata["columns"][1]["totalsRowFunction"] == "sum"
            and table_cells["A2"]["value"] == "Q2"
            and table_cells["B3"]["value"] == 30
            and table_cells["A4"]["value"] == "Q1"
            and table_cells["A5"]["value"] == "Total"
            and table_cells["B5"]["formula"] == "SUBTOTAL(109,[Revenue])"
            and table_edit["changes"][1]["preservedTotalsRow"] is True
        )
        no_totals_xlsx = root / "table-no-totals.xlsx"
        no_totals_edit = run_helper(table_xlsx, "excel", action="edit", outputPath=str(no_totals_xlsx), operations=[
            {"type": "setTableTotals", "table": "RevenueTable", "enabled": False},
        ])
        no_totals_result = run_helper(no_totals_xlsx, "excel")
        no_totals_cells = {cell["address"]: cell for cell in no_totals_result["sheets"][0]["cells"]}
        no_totals_table = no_totals_result["sheets"][0]["tables"][0]
        checks["excelDisableTotals"] = no_totals_edit["changeCount"] == 1 and no_totals_table["ref"] == "A1:B4" and no_totals_table["totalsRowCount"] == 0 and "A5" not in no_totals_cells and "B5" not in no_totals_cells

        compatibility_xlsx = root / "compatibility-table.xlsx"
        write_package(compatibility_xlsx, {
            "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Log" sheetId="1" r:id="rId1"/></sheets></workbook>""",
            "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
            "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" mc:Ignorable="x14ac xr xr2 xr3" xr:uid="{11111111-1111-1111-1111-111111111111}"><dimension ref="A1:B3"/><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Metric</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2" ht="24" customHeight="1"><c r="A2" s="2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2" s="3"><v>10</v></c></row><row r="3" ht="26" customHeight="1"><c r="A3" s="2" t="inlineStr"><is><t>Q2</t></is></c><c r="B3" s="3"><f>B2+1</f><v>11</v></c></row></sheetData><conditionalFormatting sqref="A2:B3"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" sqref="B2:B3"><formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations><hyperlinks><hyperlink ref="A2" r:id="rId2"/><hyperlink ref="A3" r:id="rId3"/></hyperlinks><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>""",
            "xl/worksheets/_rels/sheet1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/q1" TargetMode="External"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/q2" TargetMode="External"/></Relationships>""",
            "xl/tables/table1.xml": """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" mc:Ignorable="xr xr3" xr:uid="{22222222-2222-2222-2222-222222222222}" id="1" name="CompatibilityTable" displayName="CompatibilityTable" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Metric"/><tableColumn id="2" name="Value"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>""",
        })
        compatibility_output = root / "compatibility-table-edited.xlsx"
        compatibility_edit = run_helper(compatibility_xlsx, "excel", action="edit", outputPath=str(compatibility_output), operations=[
            {"type": "addTableRow", "table": "CompatibilityTable", "values": [["Q3", 30]]},
        ])
        compatibility_validation = run_helper(compatibility_output, "excel", action="validate")
        with zipfile.ZipFile(compatibility_output, "r") as package:
            sheet_root = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
            table_xml = package.read("xl/tables/table1.xml").decode("utf-8")
            sheet_xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
            spreadsheet_ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            new_row = next(row for row in sheet_root.find(spreadsheet_ns + "sheetData") if row.attrib.get("r") == "4")
            new_cells = {cell.attrib.get("r"): cell for cell in new_row}
            conditional = sheet_root.find(spreadsheet_ns + "conditionalFormatting")
            validation = sheet_root.find(spreadsheet_ns + "dataValidations").find(spreadsheet_ns + "dataValidation")
        checks["excelCompatibilityNamespacePreservation"] = (
            compatibility_validation["valid"]
            and compatibility_validation["package"]["markupCompatibilityErrors"] == []
            and all(("xmlns:%s=" % prefix) in sheet_xml for prefix in ("x14ac", "xr", "xr2", "xr3"))
            and all(("xmlns:%s=" % prefix) in table_xml for prefix in ("xr", "xr3"))
        )
        checks["excelAppendCopiesRowMetadata"] = (
            compatibility_edit["changes"][0]["copiedFormatFromRow"] == 3
            and compatibility_edit["changes"][0]["expandedConditionalFormatting"] == 1
            and compatibility_edit["changes"][0]["expandedDataValidations"] == 1
            and new_row.attrib.get("ht") == "26"
            and new_row.attrib.get("customHeight") == "1"
            and new_cells["A4"].attrib.get("s") == "2"
            and new_cells["B4"].attrib.get("s") == "3"
            and conditional.attrib.get("sqref") == "A2:B4"
            and validation.attrib.get("sqref") == "B2:B4"
        )

        compatibility_deleted = root / "compatibility-table-deleted.xlsx"
        compatibility_delete = run_helper(compatibility_output, "excel", action="edit", outputPath=str(compatibility_deleted), operations=[
            {"type": "deleteTableRow", "table": "CompatibilityTable", "index": 0},
        ])
        compatibility_delete_validation = run_helper(compatibility_deleted, "excel", action="validate")
        compatibility_delete_result = run_helper(compatibility_deleted, "excel")
        deleted_cells = {cell["address"]: cell for cell in compatibility_delete_result["sheets"][0]["cells"]}
        deleted_table = compatibility_delete_result["sheets"][0]["tables"][0]
        with zipfile.ZipFile(compatibility_deleted, "r") as package:
            deleted_sheet = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
            deleted_relationships = ET.fromstring(package.read("xl/worksheets/_rels/sheet1.xml.rels"))
            deleted_hyperlinks = deleted_sheet.find(spreadsheet_ns + "hyperlinks")
            deleted_conditional = deleted_sheet.find(spreadsheet_ns + "conditionalFormatting")
            deleted_validation = deleted_sheet.find(spreadsheet_ns + "dataValidations").find(spreadsheet_ns + "dataValidation")
            deleted_dimension = deleted_sheet.find(spreadsheet_ns + "dimension")
            deleted_row = next(row for row in deleted_sheet.find(spreadsheet_ns + "sheetData") if row.attrib.get("r") == "2")
            relationship_ids = {entry.attrib.get("Id") for entry in list(deleted_relationships)}
        delete_change = compatibility_delete["changes"][0]
        checks["excelDeleteTableRowCompactsSafely"] = (
            compatibility_delete_validation["valid"]
            and compatibility_delete_validation["package"]["markupCompatibilityErrors"] == []
            and deleted_table["ref"] == "A1:B3"
            and deleted_cells["A2"]["value"] == "Q2"
            and deleted_cells["B2"]["formula"] == "B1+1"
            and deleted_cells["A3"]["value"] == "Q3"
            and "A4" not in deleted_cells and "B4" not in deleted_cells
            and deleted_row.attrib.get("ht") == "26"
            and [cell.attrib.get("r") for cell in list(deleted_row)] == ["A2", "B2"]
            and deleted_conditional.attrib.get("sqref") == "A2:B3"
            and deleted_validation.attrib.get("sqref") == "B2:B3"
            and deleted_dimension.attrib.get("ref") == "A1:B3"
            and len(list(deleted_hyperlinks)) == 1
            and list(deleted_hyperlinks)[0].attrib.get("ref") == "A2"
            and "rId2" not in relationship_ids and "rId3" in relationship_ids
            and delete_change["beforeRef"] == "A1:B4"
            and delete_change["afterRef"] == "A1:B3"
            and delete_change["formulasTranslated"] == 1
            and delete_change["shrunkConditionalFormatting"] == 1
            and delete_change["shrunkDataValidations"] == 1
            and delete_change["removedHyperlinks"] == 1
            and delete_change["shiftedHyperlinks"] == 1
            and delete_change["removedHyperlinkRelationships"] == 1
        )

        shared_formula_xlsx = root / "shared-formula-table.xlsx"
        write_package(shared_formula_xlsx, {
            "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Shared" sheetId="1" r:id="rId1"/></sheets></workbook>""",
            "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
            "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Metric</t></is></c><c r="B1" t="inlineStr"><is><t>Fit</t></is></c></row><row r="2" ht="22" customHeight="1"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><f t="shared" ref="B2:B4" si="9">A2&amp;\"-fit\"</f><v>101</v></c><c r="C2" s="5"/></row><row r="3" ht="24" customHeight="1"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><f t="shared" si="9"/><v>202</v></c><c r="C3" s="5"/></row><row r="4" ht="26" customHeight="1"><c r="A4" t="inlineStr"><is><t>Q3</t></is></c><c r="B4"><f t="shared" si="9"/><v>303</v></c><c r="C4" s="5"/></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>""",
            "xl/worksheets/_rels/sheet1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>""",
            "xl/tables/table1.xml": """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SharedTable" displayName="SharedTable" ref="A1:B4"><autoFilter ref="A1:B4"/><tableColumns count="2"><tableColumn id="1" name="Metric"/><tableColumn id="2" name="Fit"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>""",
        })
        shared_formula_deleted = root / "shared-formula-table-deleted.xlsx"
        shared_formula_edit = run_helper(shared_formula_xlsx, "excel", action="edit", outputPath=str(shared_formula_deleted), operations=[
            {"type": "deleteTableRow", "table": "SharedTable", "index": 1},
        ])
        shared_formula_validation = run_helper(shared_formula_deleted, "excel", action="validate")
        with zipfile.ZipFile(shared_formula_deleted, "r") as package:
            shared_sheet = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
            shared_rows = {row.attrib["r"]: row for row in shared_sheet.find(spreadsheet_ns + "sheetData")}
            shared_master = next(cell for cell in shared_rows["2"] if cell.attrib.get("r") == "B2")
            shared_follower = next(cell for cell in shared_rows["3"] if cell.attrib.get("r") == "B3")
            shared_master_formula = shared_master.find(spreadsheet_ns + "f")
            shared_follower_formula = shared_follower.find(spreadsheet_ns + "f")
            shared_follower_cache = shared_follower.find(spreadsheet_ns + "v")
        shared_change = shared_formula_edit["changes"][0]
        checks["excelDeleteTableRowPreservesSharedFormulas"] = (
            shared_formula_validation["valid"]
            and shared_master_formula.attrib == {"t": "shared", "ref": "B2:B3", "si": "9"}
            and shared_master_formula.text == 'A2&"-fit"'
            and shared_follower_formula.attrib == {"t": "shared", "si": "9"}
            and shared_follower_formula.text is None
            and shared_follower_cache.text == "303"
            and shared_rows["3"].attrib.get("ht") == "26"
            and [cell.attrib.get("r") for cell in list(shared_rows["3"])] == ["A3", "B3", "C3"]
            and shared_change["sharedFormulaRefsShrunk"] == 1
            and shared_change["sharedFormulaCellsMoved"] == 1
            and shared_change["formulasTranslated"] == 0
            and shared_change["shiftedRowAttributes"] == 1
        )

        bad_compatibility = root / "bad-compatibility.xlsx"
        write_package(bad_compatibility, {
            "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>""",
            "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
            "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr3"><sheetData/></worksheet>""",
        })
        bad_compatibility_result = subprocess.run(
            [sys.executable, str(HELPER)],
            input=json.dumps({"action": "validate", "inputPath": str(bad_compatibility), "kind": "excel"}),
            text=True,
            capture_output=True,
            check=False,
        )
        checks["undeclaredIgnorablePrefixRejected"] = bad_compatibility_result.returncode != 0 and "invalid markup-compatibility declarations" in bad_compatibility_result.stdout

        pptx = fixtures["powerpoint"]
        powerpoint = run_helper(pptx, "powerpoint", searchText="PowerPoint")
        checks["powerpoint"] = powerpoint["slideCount"] == 1 and powerpoint["slides"][0]["shapes"][0]["text"] == "Hello PowerPoint" and powerpoint["search"]["matchCount"] == 1
        edited_pptx = root / "edited.pptx"
        ppt_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(edited_pptx), operations=[{"type": "replaceText", "slideIndex": 0, "shapeId": "2", "find": "Hello", "replace": "Edited"}])
        edited_powerpoint = run_helper(edited_pptx, "powerpoint")
        checks["powerpointEdit"] = ppt_edit["changeCount"] == 1 and edited_powerpoint["slides"][0]["shapes"][0]["text"] == "Edited PowerPoint"
        rich_pptx = root / "rich.pptx"
        rich_ppt_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(rich_pptx), operations=[
            {"type": "setShapeText", "slideIndex": 0, "shapeId": "2", "text": "Native PowerPoint"},
            {"type": "setShapeGeometry", "slideIndex": 0, "shapeId": "2", "x": 10, "y": 20, "width": 30, "height": 40},
            {"type": "setTableCell", "slideIndex": 0, "shapeId": "4", "rowIndex": 0, "columnIndex": 0, "text": "Updated table value"},
            {"type": "setNotes", "slideIndex": 0, "text": "Updated speaker notes"},
        ])
        rich_powerpoint = run_helper(rich_pptx, "powerpoint")
        shape = rich_powerpoint["slides"][0]["shapes"][0]
        table_shape = next(entry for entry in rich_powerpoint["slides"][0]["shapes"] if entry["id"] == "4")
        checks["powerpointStructuredEdits"] = rich_ppt_edit["changeCount"] == 4 and shape["text"] == "Native PowerPoint" and all(shape["geometry"][key] == value for key, value in {"x": 10, "y": 20, "width": 30, "height": 40}.items()) and table_shape["table"]["rows"] == [["Updated table value"]] and rich_powerpoint["slides"][0]["notes"] == "Updated speaker notes"
        empty_pptx = root / "empty-table-no-notes.pptx"
        with zipfile.ZipFile(pptx, "r") as source, zipfile.ZipFile(empty_pptx, "w", zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                if info.filename == "ppt/notesSlides/notesSlide1.xml":
                    continue
                payload = source.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    payload = payload.replace(b'<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>', b"")
                elif info.filename == "ppt/slides/_rels/slide1.xml.rels":
                    payload = payload.replace(b'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>', b"")
                elif info.filename == "ppt/slides/slide1.xml":
                    payload = payload.replace(b"<a:r><a:t>Old table value</a:t></a:r>", b"")
                target.writestr(info, payload)
        repaired_pptx = root / "empty-table-notes-repaired.pptx"
        repaired_edit = run_helper(empty_pptx, "powerpoint", action="edit", outputPath=str(repaired_pptx), operations=[
            {"type": "setTableCell", "slideIndex": 0, "shapeId": "4", "rowIndex": 0, "columnIndex": 0, "text": "Created table text"},
            {"type": "setNotes", "slideIndex": 0, "text": "Created speaker notes"},
        ])
        repaired_powerpoint = run_helper(repaired_pptx, "powerpoint")
        repaired_table = next(entry for entry in repaired_powerpoint["slides"][0]["shapes"] if entry["id"] == "4")
        checks["powerpointCreatesMissingTextAndNotes"] = (
            repaired_edit["changeCount"] == 2
            and repaired_table["table"]["rows"] == [["Created table text"]]
            and repaired_powerpoint["slides"][0]["notes"] == "Created speaker notes"
        )
        duplicated_pptx = root / "slides-duplicated.pptx"
        duplicate_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(duplicated_pptx), operations=[{"type": "duplicateSlide", "slideIndex": 0}])
        duplicated_powerpoint = run_helper(duplicated_pptx, "powerpoint")
        distinguished_pptx = root / "slides-distinguished.pptx"
        distinguish_edit = run_helper(duplicated_pptx, "powerpoint", action="edit", outputPath=str(distinguished_pptx), operations=[{"type": "setShapeText", "slideIndex": 1, "shapeId": "2", "text": "Duplicate copy"}])
        moved_pptx = root / "slides-moved.pptx"
        move_edit = run_helper(distinguished_pptx, "powerpoint", action="edit", outputPath=str(moved_pptx), operations=[{"type": "moveSlide", "slideIndex": 1, "toIndex": 0}])
        moved_powerpoint = run_helper(moved_pptx, "powerpoint")
        moved_text = [slide["shapes"][0]["text"] for slide in moved_powerpoint["slides"]]
        deleted_pptx = root / "slides-deleted.pptx"
        delete_edit = run_helper(moved_pptx, "powerpoint", action="edit", outputPath=str(deleted_pptx), operations=[{"type": "deleteSlide", "slideIndex": 1}])
        deleted_powerpoint = run_helper(deleted_pptx, "powerpoint")
        checks["powerpointSlideLifecycle"] = (
            duplicate_edit["changeCount"] == 1
            and duplicated_powerpoint["slideCount"] == 2
            and distinguish_edit["changeCount"] == 1
            and move_edit["changeCount"] == 1
            and moved_text == ["Duplicate copy", "Hello PowerPoint"]
            and delete_edit["changeCount"] == 1
            and deleted_powerpoint["slideCount"] == 1
            and deleted_powerpoint["slides"][0]["shapes"][0]["text"] == "Duplicate copy"
        )

        media_pptx = root / "media.pptx"
        media_content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"""
        media_slide = """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Delete me"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old text</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 1" descr="replace target"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"""
        write_package(media_pptx, {
            "_rels/.rels": root_rels("ppt/presentation.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "ppt/presentation.xml": """<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>""",
            "ppt/_rels/presentation.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""",
            "ppt/slides/slide1.xml": media_slide,
            "ppt/slides/_rels/slide1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.gif"/></Relationships>""",
            "ppt/media/image1.gif": b"GIF89a",
        }, media_content_types)
        native_pptx = root / "native-operations.pptx"
        replacement_png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        native_edit = run_helper(media_pptx, "powerpoint", action="edit", outputPath=str(native_pptx), operations=[
            {"type": "addTextBox", "slideIndex": 0, "shapeId": 4, "name": "Added box", "text": "Styled text", "x": 1000, "y": 2000, "width": 3000, "height": 4000},
            {"type": "setTextStyle", "slideIndex": 0, "shapeId": 4, "fontFamily": "Aptos", "fontSize": 18, "bold": True, "italic": True, "underline": True, "color": "12AB34"},
            {"type": "deleteShape", "slideIndex": 0, "shapeId": 2},
            {"type": "replaceImage", "slideIndex": 0, "shapeId": 3, "base64": replacement_png, "contentType": "image/png"},
        ])
        native_result = run_helper(native_pptx, "powerpoint")
        native_shapes = {shape["id"]: shape for shape in native_result["slides"][0]["shapes"]}
        styled = native_shapes["4"]["paragraphs"][0]["runs"][0]["properties"]
        with zipfile.ZipFile(native_pptx, "r") as package:
            slide_rels = package.read("ppt/slides/_rels/slide1.xml.rels").decode("utf-8")
            slide_xml = package.read("ppt/slides/slide1.xml").decode("utf-8")
            content_types = package.read("[Content_Types].xml").decode("utf-8")
            new_media = [name for name in package.namelist() if name.startswith("ppt/media/image") and name.endswith(".png")]
            preserved_source = package.read("ppt/media/image1.gif") == b"GIF89a"
        checks["powerpointNativeShapeAndImageEdits"] = (native_edit["changeCount"] == 4 and "2" not in native_shapes and native_shapes["4"]["text"] == "Styled text" and styled.get("sz") == "1800" and styled.get("b") == "1" and styled.get("i") == "1" and styled.get("u") == "sng" and "Aptos" in slide_xml and "12AB34" in slide_xml and len(new_media) == 1 and "../media/" + Path(new_media[0]).name in slide_rels and "image/png" in content_types and preserved_source)

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

    advertised_openxml_operations = production_openxml_operations()
    real_fixture_source = (ROOT / "scripts" / "office-real-fixture-test.py").read_text()
    checks["richFixtureContractDeclared"] = all(
        ('"type": "%s"' % operation) in real_fixture_source
        for operations in RICH_REAL_FIXTURE_OPERATIONS.values() for operation in operations
    )
    checks["advertisedOpenXmlOperationCounts"] = {
        kind: len(operations) for kind, operations in advertised_openxml_operations.items()
    } == EXPECTED_OPENXML_OPERATION_COUNTS
    coverage = {
        kind: {
            "covered": sorted(COVERED_OPENXML_OPERATIONS[kind] | RICH_REAL_FIXTURE_OPERATIONS[kind]),
            "expected": sorted(expected),
            "coveredCount": len(COVERED_OPENXML_OPERATIONS[kind] | RICH_REAL_FIXTURE_OPERATIONS[kind]),
            "expectedCount": len(expected),
        }
        for kind, expected in advertised_openxml_operations.items()
    }
    checks["advertisedOpenXmlCoverage"] = all(set(entry["covered"]) == set(entry["expected"]) for entry in coverage.values())
    ok = all(checks.values())
    print(json.dumps({"ok": ok, "checks": checks, "operationCoverage": coverage}, indent=2))
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    fixture_arg = next((arg for arg in sys.argv[1:] if arg.startswith("--emit-fixtures=")), None)
    if fixture_arg:
        fixtures = emit_fixtures(Path(fixture_arg.split("=", 1)[1]).resolve())
        print(json.dumps({key: str(value) for key, value in fixtures.items()}))
    else:
        main()
