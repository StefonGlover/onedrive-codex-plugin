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
        "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>""",
        "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><f>SUM(1,2)</f><v>3</v></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts><pivotTableParts count="1"><pivotTablePart r:id="rId3"/></pivotTableParts><drawing r:id="rId2"/></worksheet>""",
        "xl/worksheets/_rels/sheet1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>""",
        "xl/tables/table1.xml": """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="RevenueTable" displayName="RevenueTable" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="Metric"/><tableColumn id="2" name="Revenue"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>""",
        "xl/drawings/drawing1.xml": """<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:graphicFrame><c:chart r:id="rId1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>""",
        "xl/drawings/_rels/drawing1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>""",
        "xl/charts/chart1.xml": """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue Chart</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>""",
        "xl/pivotTables/pivotTable1.xml": """<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="RevenuePivot" cacheId="1"><location ref="D1:F5"/><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="1"/></dataFields></pivotTableDefinition>""",
        "xl/styles.xml": """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>""",
        "xl/calcChain.xml": """<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>""",
    })
    pptx = root / "sample.pptx"
    ppt_content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"""
    write_package(pptx, {
        "_rels/.rels": root_rels("ppt/presentation.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        "ppt/presentation.xml": """<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>""",
        "ppt/_rels/presentation.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""",
        "ppt/slides/slide1.xml": """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello PowerPoint</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>""",
        "ppt/slides/_rels/slide1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.gif"/></Relationships>""",
        "ppt/media/image1.gif": b"GIF89a",
    }, ppt_content_types)
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

        xlsx = root / "sample.xlsx"
        write_package(xlsx, {
            "_rels/.rels": root_rels("xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Total">Data!$A$1</definedName></definedNames></workbook>""",
            "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>""",
            "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><f>SUM(1,2)</f><v>3</v></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts><pivotTableParts count="1"><pivotTablePart r:id="rId3"/></pivotTableParts><drawing r:id="rId2"/></worksheet>""",
            "xl/worksheets/_rels/sheet1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>""",
            "xl/tables/table1.xml": """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="RevenueTable" displayName="RevenueTable" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="Metric"/><tableColumn id="2" name="Revenue"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>""",
            "xl/drawings/drawing1.xml": """<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:graphicFrame><c:chart r:id="rId1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>""",
            "xl/drawings/_rels/drawing1.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>""",
            "xl/charts/chart1.xml": """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue Chart</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>""",
            "xl/pivotTables/pivotTable1.xml": """<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="RevenuePivot" cacheId="1"><location ref="D1:F5"/><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="1"/></dataFields></pivotTableDefinition>""",
            "xl/styles.xml": """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>""",
            "xl/calcChain.xml": """<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>""",
        })
        excel = run_helper(xlsx, "excel", searchText="Revenue")
        cells = excel["sheets"][0]["cells"]
        checks["excel"] = excel["sheetCount"] == 1 and cells[0]["value"] == "Revenue" and excel["search"]["matchCount"] == 1 and excel["tableCount"] == 1 and excel["sheets"][0]["tables"][0]["displayName"] == "RevenueTable" and excel["chartCount"] == 1 and excel["sheets"][0]["charts"][0]["title"] == "Revenue Chart" and excel["pivotCount"] == 1 and excel["sheets"][0]["pivots"][0]["name"] == "RevenuePivot"
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

        charts_xlsx = root / "chart-operations.xlsx"
        chart_edit = run_helper(xlsx, "excel", action="edit", outputPath=str(charts_xlsx), operations=[
            {"type": "createChart", "sheet": "Data", "chartType": "ColumnClustered", "sourceData": "A1:B2", "name": "Consumer Chart", "titleText": "Consumer Revenue", "left": 12, "top": 18, "width": 360, "height": 220},
            {"type": "updateChart", "sheet": "Data", "chart": "Consumer Chart", "chartType": "Line", "sourceData": "A1:B2", "name": "Consumer Trend", "titleText": "Revenue Trend", "left": 20, "top": 24, "width": 400, "height": 240},
        ])
        chart_result = run_helper(charts_xlsx, "excel")
        created_chart = next(chart for chart in chart_result["sheets"][0]["charts"] if chart.get("name") == "Consumer Trend")
        checks["excelChartLifecycle"] = chart_edit["changeCount"] == 2 and created_chart["type"] == "lineChart" and created_chart["title"] == "Revenue Trend" and created_chart["seriesCount"] == 1 and round(created_chart["geometry"]["width"]) == 400 and round(created_chart["geometry"]["height"]) == 240

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

        pptx = root / "sample.pptx"
        write_package(pptx, {
            "_rels/.rels": root_rels("ppt/presentation.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            "ppt/presentation.xml": """<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>""",
            "ppt/_rels/presentation.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""",
            "ppt/slides/slide1.xml": """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello PowerPoint</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>""",
        })
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
        ])
        rich_powerpoint = run_helper(rich_pptx, "powerpoint")
        shape = rich_powerpoint["slides"][0]["shapes"][0]
        checks["powerpointStructuredEdits"] = rich_ppt_edit["changeCount"] == 2 and shape["text"] == "Native PowerPoint" and all(shape["geometry"][key] == value for key, value in {"x": 10, "y": 20, "width": 30, "height": 40}.items())
        slides_pptx = root / "slides.pptx"
        slides_edit = run_helper(pptx, "powerpoint", action="edit", outputPath=str(slides_pptx), operations=[{"type": "duplicateSlide", "slideIndex": 0}, {"type": "moveSlide", "slideIndex": 1, "toIndex": 0}])
        slides_powerpoint = run_helper(slides_pptx, "powerpoint")
        checks["powerpointSlideLifecycle"] = slides_edit["changeCount"] == 2 and slides_powerpoint["slideCount"] == 2 and all(slide["shapes"][0]["text"] == "Hello PowerPoint" for slide in slides_powerpoint["slides"])

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
