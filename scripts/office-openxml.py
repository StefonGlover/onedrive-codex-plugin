#!/usr/bin/env python3
"""Dependency-free Office Open XML inspection and validation helper.

The MCP server communicates with this process using one JSON request on stdin
and receives one JSON response on stdout.  The helper intentionally edits only
explicit package parts and copies every untouched ZIP member byte-for-byte.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlparse
from xml.etree import ElementTree as ET


MAX_PACKAGE_ENTRIES = 10_000
MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_COMPRESSION_RATIO = 250
MAX_TEXT_CHARS = 2_000_000

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}

for _prefix, _uri in NS.items():
    if _prefix not in {"ct", "pr"}:
        ET.register_namespace(_prefix, _uri)


class OfficePackageError(Exception):
    pass


def q(prefix: str, local: str) -> str:
    return "{%s}%s" % (NS[prefix], local)


def read_request() -> Dict[str, Any]:
    raw = sys.stdin.read(MAX_TEXT_CHARS + 1)
    if len(raw) > MAX_TEXT_CHARS:
        raise OfficePackageError("Office helper request is too large.")
    try:
        value = json.loads(raw or "{}")
    except json.JSONDecodeError as error:
        raise OfficePackageError("Invalid JSON request: %s" % error) from error
    if not isinstance(value, dict):
        raise OfficePackageError("Office helper request must be a JSON object.")
    return value


def safe_member_name(name: str) -> bool:
    path = PurePosixPath(name)
    return not path.is_absolute() and ".." not in path.parts and "\\" not in name


def package_inventory(package: zipfile.ZipFile) -> Dict[str, Any]:
    infos = package.infolist()
    if len(infos) > MAX_PACKAGE_ENTRIES:
        raise OfficePackageError("Office package contains too many ZIP entries.")
    names_in_order = [info.filename for info in infos]
    if len(names_in_order) != len(set(names_in_order)):
        raise OfficePackageError("Office package contains duplicate ZIP member names.")
    total_uncompressed = sum(info.file_size for info in infos)
    if total_uncompressed > MAX_UNCOMPRESSED_BYTES:
        raise OfficePackageError("Office package expands beyond the configured safety limit.")
    unsafe = [info.filename for info in infos if not safe_member_name(info.filename)]
    if unsafe:
        raise OfficePackageError("Office package contains unsafe ZIP paths.")
    suspicious = []
    encrypted = []
    for info in infos:
        if info.flag_bits & 0x1:
            encrypted.append(info.filename)
        compressed = max(info.compress_size, 1)
        if info.file_size > 1024 * 1024 and info.file_size / compressed > MAX_COMPRESSION_RATIO:
            suspicious.append(info.filename)
    if encrypted:
        raise OfficePackageError("Encrypted Office packages are not supported.")
    if suspicious:
        raise OfficePackageError("Office package contains suspiciously compressed entries.")
    names = {info.filename for info in infos}
    return {
        "entryCount": len(infos),
        "compressedBytes": sum(info.compress_size for info in infos),
        "uncompressedBytes": total_uncompressed,
        "names": names,
        "hasMacros": any(name.lower().endswith("vbaproject.bin") for name in names),
        "hasDigitalSignatures": any(
            name.lower().startswith("_xmlsignatures/") or "digitalsignature" in name.lower()
            for name in names
        ),
    }


def xml_root(package: zipfile.ZipFile, name: str) -> ET.Element:
    try:
        payload = package.read(name)
    except KeyError as error:
        raise OfficePackageError("Required Office package part is missing: %s" % name) from error
    try:
        return ET.fromstring(payload)
    except ET.ParseError as error:
        raise OfficePackageError("Invalid XML in Office package part %s: %s" % (name, error)) from error


def detect_kind(package: zipfile.ZipFile, names: Iterable[str]) -> str:
    name_set = set(names)
    if "word/document.xml" in name_set:
        return "word"
    if "xl/workbook.xml" in name_set:
        return "excel"
    if "ppt/presentation.xml" in name_set:
        return "powerpoint"
    raise OfficePackageError("The file is not a supported .docx, .xlsx, or .pptx Open XML package.")


def rels_for_part(package: zipfile.ZipFile, part_name: str) -> Dict[str, Dict[str, str]]:
    directory = posixpath.dirname(part_name)
    rel_name = posixpath.join(directory, "_rels", posixpath.basename(part_name) + ".rels")
    if rel_name not in package.namelist():
        return {}
    root = xml_root(package, rel_name)
    result: Dict[str, Dict[str, str]] = {}
    for rel in root.findall(q("pr", "Relationship")):
        rel_id = rel.attrib.get("Id")
        if rel_id:
            result[rel_id] = dict(rel.attrib)
    return result


def resolved_relationship_target(part_name: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part_name), target))


def validate_relationships(package: zipfile.ZipFile, names: Iterable[str]) -> List[Dict[str, str]]:
    name_set = set(names)
    errors: List[Dict[str, str]] = []
    for rel_name in sorted(name for name in name_set if name.endswith(".rels")):
        root = xml_root(package, rel_name)
        if rel_name == "_rels/.rels":
            source_part = ""
        else:
            directory = posixpath.dirname(posixpath.dirname(rel_name))
            base = posixpath.basename(rel_name)[:-5]
            source_part = posixpath.join(directory, base)
        for rel in root.findall(q("pr", "Relationship")):
            if rel.attrib.get("TargetMode") == "External":
                continue
            target = rel.attrib.get("Target")
            if not target:
                errors.append({"part": rel_name, "error": "relationship target is empty"})
                continue
            resolved = resolved_relationship_target(source_part, target)
            if resolved not in name_set:
                errors.append({"part": rel_name, "target": resolved, "error": "relationship target is missing"})
    return errors


def validate_content_types(package: zipfile.ZipFile, names: Iterable[str]) -> List[Dict[str, str]]:
    name_set = set(names)
    if "[Content_Types].xml" not in name_set:
        return [{"part": "[Content_Types].xml", "error": "content types manifest is missing"}]
    root = xml_root(package, "[Content_Types].xml")
    defaults = {entry.attrib.get("Extension", "").lower() for entry in root.findall(q("ct", "Default"))}
    overrides = {entry.attrib.get("PartName", "").lstrip("/") for entry in root.findall(q("ct", "Override"))}
    errors = []
    for name in sorted(name_set):
        if name == "[Content_Types].xml" or name.endswith("/"):
            continue
        extension = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if name not in overrides and extension not in defaults:
            errors.append({"part": name, "error": "part has no Default or Override content type"})
    return errors


def text_of(element: ET.Element, tags: Tuple[str, ...]) -> str:
    return "".join(node.text or "" for node in element.iter() if node.tag in tags)


def sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def package_fingerprint(package: zipfile.ZipFile) -> str:
    digest = hashlib.sha256()
    for name in sorted(package.namelist()):
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(package.read(name))
        digest.update(b"\0")
    return digest.hexdigest()


def word_paragraph(element: ET.Element, index: int, part: str) -> Dict[str, Any]:
    ppr = element.find(q("w", "pPr"))
    style = None
    if ppr is not None:
        style_node = ppr.find(q("w", "pStyle"))
        if style_node is not None:
            style = style_node.attrib.get(q("w", "val"))
    text = text_of(element, (q("w", "t"), q("w", "tab"), q("w", "br")))
    runs = []
    for run in element.iter(q("w", "r")):
        props = run.find(q("w", "rPr"))
        runs.append({
            "text": text_of(run, (q("w", "t"), q("w", "tab"), q("w", "br"))),
            "style": props.find(q("w", "rStyle")).attrib.get(q("w", "val")) if props is not None and props.find(q("w", "rStyle")) is not None else None,
            "bold": props.find(q("w", "b")) is not None if props is not None else False,
            "italic": props.find(q("w", "i")) is not None if props is not None else False,
            "underline": props.find(q("w", "u")).attrib.get(q("w", "val"), "single") if props is not None and props.find(q("w", "u")) is not None else None,
        })
    return {
        "index": index,
        "part": part,
        "text": text,
        "style": style,
        "runs": runs,
        "numbering": ({"numId": ppr.find(q("w", "numPr")).find(q("w", "numId")).attrib.get(q("w", "val"))} if ppr is not None and ppr.find(q("w", "numPr")) is not None and ppr.find(q("w", "numPr")).find(q("w", "numId")) is not None else None),
        "hasTrackedChanges": any(node.tag in {q("w", "ins"), q("w", "del")} for node in element.iter()),
    }


def inspect_word(package: zipfile.ZipFile, request: Dict[str, Any]) -> Dict[str, Any]:
    names = set(package.namelist())
    parts = ["word/document.xml"]
    if request.get("includeHeadersFooters", True):
        parts.extend(sorted(name for name in names if re.fullmatch(r"word/(header|footer)\d+\.xml", name)))
    paragraphs: List[Dict[str, Any]] = []
    tables: List[Dict[str, Any]] = []
    content_controls: List[Dict[str, Any]] = []
    max_paragraphs = max(1, min(int(request.get("maxParagraphs", 2000)), 10_000))
    for part in parts:
        root = xml_root(package, part)
        for paragraph in root.iter(q("w", "p")):
            if len(paragraphs) >= max_paragraphs:
                break
            paragraphs.append(word_paragraph(paragraph, len(paragraphs), part))
        for table in root.iter(q("w", "tbl")):
            rows = []
            for row in table.findall(q("w", "tr")):
                rows.append([text_of(cell, (q("w", "t"),)) for cell in row.findall(q("w", "tc"))])
            tables.append({"index": len(tables), "part": part, "rows": rows})
        for sdt in root.iter(q("w", "sdt")):
            props = sdt.find(q("w", "sdtPr"))
            tag = title = control_id = None
            if props is not None:
                tag_node = props.find(q("w", "tag"))
                title_node = props.find(q("w", "alias"))
                id_node = props.find(q("w", "id"))
                tag = tag_node.attrib.get(q("w", "val")) if tag_node is not None else None
                title = title_node.attrib.get(q("w", "val")) if title_node is not None else None
                control_id = id_node.attrib.get(q("w", "val")) if id_node is not None else None
            content_controls.append({
                "index": len(content_controls),
                "part": part,
                "id": control_id,
                "tag": tag,
                "title": title,
                "text": text_of(sdt, (q("w", "t"),)),
            })
    comments = []
    if "word/comments.xml" in names:
        comments_root = xml_root(package, "word/comments.xml")
        for comment in comments_root.findall(q("w", "comment")):
            comments.append({
                "id": comment.attrib.get(q("w", "id")),
                "author": comment.attrib.get(q("w", "author")),
                "date": comment.attrib.get(q("w", "date")),
                "text": text_of(comment, (q("w", "t"),)),
            })
    return {
        "kind": "word",
        "paragraphs": paragraphs,
        "paragraphCount": len(paragraphs),
        "tables": tables,
        "tableCount": len(tables),
        "contentControls": content_controls,
        "contentControlCount": len(content_controls),
        "comments": comments,
        "commentCount": len(comments),
        "truncated": len(paragraphs) >= max_paragraphs,
    }


def excel_shared_strings(package: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    root = xml_root(package, "xl/sharedStrings.xml")
    return [text_of(item, (q("s", "t"),)) for item in root.findall(q("s", "si"))]


def excel_cell_value(cell: ET.Element, shared: List[str]) -> Any:
    value_node = cell.find(q("s", "v"))
    inline = cell.find(q("s", "is"))
    raw = value_node.text if value_node is not None else None
    cell_type = cell.attrib.get("t")
    if cell_type == "s" and raw is not None:
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "inlineStr" and inline is not None:
        return text_of(inline, (q("s", "t"),))
    if cell_type == "b" and raw is not None:
        return raw == "1"
    if raw is None:
        return None
    try:
        number = float(raw)
        return int(number) if number.is_integer() else number
    except ValueError:
        return raw


def excel_tables_for_sheet(package: zipfile.ZipFile, sheet_part: str) -> List[Dict[str, Any]]:
    relationships = rels_for_part(package, sheet_part)
    tables = []
    for relationship_id, relationship in relationships.items():
        if not relationship.get("Type", "").endswith("/table"):
            continue
        target = relationship.get("Target", "")
        part = resolved_relationship_target(sheet_part, target) if target else None
        if not part or part not in package.namelist():
            continue
        root = xml_root(package, part)
        columns_node = root.find(q("s", "tableColumns"))
        style = root.find(q("s", "tableStyleInfo"))
        tables.append({
            "id": root.attrib.get("id"),
            "name": root.attrib.get("name"),
            "displayName": root.attrib.get("displayName"),
            "ref": root.attrib.get("ref"),
            "headerRowCount": int(root.attrib.get("headerRowCount", "1")),
            "totalsRowCount": int(root.attrib.get("totalsRowCount", "0")),
            "relationshipId": relationship_id,
            "part": part,
            "columns": [
                {
                    "id": column.attrib.get("id"),
                    "name": column.attrib.get("name"),
                    "totalsRowFunction": column.attrib.get("totalsRowFunction"),
                    "totalsRowLabel": column.attrib.get("totalsRowLabel"),
                }
                for column in list(columns_node or [])
            ],
            "style": dict(style.attrib) if style is not None else None,
        })
    return tables


def excel_charts_for_sheet(package: zipfile.ZipFile, sheet_part: str) -> List[Dict[str, Any]]:
    charts = []
    for relationship in rels_for_part(package, sheet_part).values():
        if not relationship.get("Type", "").endswith("/drawing"):
            continue
        drawing_part = resolved_relationship_target(sheet_part, relationship.get("Target", ""))
        if drawing_part not in package.namelist():
            continue
        drawing = xml_root(package, drawing_part)
        drawing_relationships = rels_for_part(package, drawing_part)
        for chart_node in drawing.iter(q("c", "chart")):
            rel_id = chart_node.attrib.get(q("r", "id"))
            chart_rel = drawing_relationships.get(rel_id or "", {})
            chart_part = resolved_relationship_target(drawing_part, chart_rel.get("Target", "")) if chart_rel else None
            if not chart_part or chart_part not in package.namelist():
                continue
            root = xml_root(package, chart_part)
            plot = root.find(".//" + q("c", "plotArea"))
            chart_type = next((node.tag.split("}")[-1] for node in list(plot or []) if node.tag.endswith("Chart")), None)
            title = text_of(root.find(".//" + q("c", "title")) or ET.Element("empty"), (q("a", "t"), q("c", "v")))
            series = []
            for entry in root.findall(".//" + q("c", "ser")):
                formulas = [node.text for node in entry.iter(q("c", "f")) if node.text]
                series.append({"index": len(series), "formulas": formulas})
            frame = next((node for node in drawing.iter(q("xdr", "graphicFrame")) if chart_node in list(node.iter())), None)
            anchor = direct_parent(drawing, frame) if frame is not None else None
            non_visual = frame.find(".//" + q("xdr", "cNvPr")) if frame is not None else None
            offset = anchor.find(q("xdr", "from")) if anchor is not None else None
            extent = anchor.find(q("xdr", "ext")) if anchor is not None else None
            geometry = {
                "left": int(offset.find(q("xdr", "colOff")).text or "0") / 12700 if offset is not None and offset.find(q("xdr", "colOff")) is not None else 0,
                "top": int(offset.find(q("xdr", "rowOff")).text or "0") / 12700 if offset is not None and offset.find(q("xdr", "rowOff")) is not None else 0,
                "width": int(extent.attrib.get("cx", "0")) / 12700 if extent is not None else None,
                "height": int(extent.attrib.get("cy", "0")) / 12700 if extent is not None else None,
            }
            charts.append({"relationshipId": rel_id, "part": chart_part, "id": non_visual.attrib.get("id") if non_visual is not None else None, "name": non_visual.attrib.get("name") if non_visual is not None else None, "type": chart_type, "title": title or None, "geometry": geometry, "series": series, "seriesCount": len(series)})
    return charts


def excel_pivots_for_sheet(package: zipfile.ZipFile, sheet_part: str) -> List[Dict[str, Any]]:
    pivots = []
    for relationship_id, relationship in rels_for_part(package, sheet_part).items():
        if not relationship.get("Type", "").endswith("/pivotTable"):
            continue
        part = resolved_relationship_target(sheet_part, relationship.get("Target", ""))
        if part not in package.namelist():
            continue
        root = xml_root(package, part)
        location = root.find(q("s", "location"))
        pivots.append({
            "name": root.attrib.get("name"),
            "cacheId": root.attrib.get("cacheId"),
            "relationshipId": relationship_id,
            "part": part,
            "location": dict(location.attrib) if location is not None else None,
            "rowFieldCount": len(list(root.find(q("s", "rowFields")) or [])),
            "columnFieldCount": len(list(root.find(q("s", "colFields")) or [])),
            "pageFieldCount": len(list(root.find(q("s", "pageFields")) or [])),
            "dataFieldCount": len(list(root.find(q("s", "dataFields")) or [])),
        })
    return pivots


def excel_formula_dependencies(formula: Optional[str], current_sheet: str) -> List[Dict[str, str]]:
    if not formula:
        return []
    pattern = re.compile(r"(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?)")
    dependencies = []
    seen = set()
    for match in pattern.finditer(formula):
        sheet = (match.group(1) or match.group(2) or current_sheet).replace("''", "'")
        address = match.group(3).replace("$", "").upper()
        key = (sheet, address)
        if key not in seen:
            seen.add(key)
            dependencies.append({"sheet": sheet, "address": address})
    return dependencies


def inspect_excel(package: zipfile.ZipFile, request: Dict[str, Any]) -> Dict[str, Any]:
    workbook = xml_root(package, "xl/workbook.xml")
    relationships = rels_for_part(package, "xl/workbook.xml")
    shared = excel_shared_strings(package)
    max_cells = max(1, min(int(request.get("maxCells", 5000)), 50_000))
    max_matches = max(1, min(int(request.get("maxMatches", 100)), 10_000))
    include_cells = request.get("includeCells", True)
    include_tables = request.get("includeTables", True)
    include_charts = request.get("includeCharts", True)
    include_pivots = request.get("includePivots", True)
    include_dependencies = request.get("includeFormulaDependencies", False)
    requested_sheets = {str(name) for name in request.get("sheetNames", []) if str(name)}
    requested_range = str(request.get("address", "")).upper()
    selected_addresses = {
        address
        for row in excel_range_addresses(requested_range)
        for address in row
    } if requested_range else None
    search_text = str(request.get("searchText", ""))
    match_case = bool(request.get("matchCase", False))
    search_needle = search_text if match_case else search_text.lower()
    sheets = []
    matches = []
    total_cells = 0
    total_tables = 0
    total_charts = 0
    total_pivots = 0
    dependency_edges = []
    found_sheet_names = set()
    sheet_nodes = workbook.find(q("s", "sheets"))
    for sheet in list(sheet_nodes or []):
        sheet_name = sheet.attrib.get("name", "")
        if requested_sheets and sheet_name not in requested_sheets:
            continue
        found_sheet_names.add(sheet_name)
        rel_id = sheet.attrib.get(q("r", "id"))
        relationship = relationships.get(rel_id or "", {})
        target = relationship.get("Target", "")
        part = resolved_relationship_target("xl/workbook.xml", target) if target else None
        sheet_result: Dict[str, Any] = {
            "name": sheet_name,
            "sheetId": sheet.attrib.get("sheetId"),
            "relationshipId": rel_id,
            "state": sheet.attrib.get("state", "visible"),
            "part": part,
        }
        cells = []
        tables = excel_tables_for_sheet(package, part) if include_tables and part and part in package.namelist() else []
        charts = excel_charts_for_sheet(package, part) if include_charts and part and part in package.namelist() else []
        pivots = excel_pivots_for_sheet(package, part) if include_pivots and part and part in package.namelist() else []
        total_tables += len(tables)
        total_charts += len(charts)
        total_pivots += len(pivots)
        if (include_cells or search_text) and part and part in package.namelist():
            root = xml_root(package, part)
            for cell in root.iter(q("s", "c")):
                if total_cells >= max_cells:
                    break
                address = cell.attrib.get("r")
                if selected_addresses is not None and address not in selected_addresses:
                    continue
                formula = cell.find(q("s", "f"))
                value = excel_cell_value(cell, shared)
                formula_text = formula.text if formula is not None else None
                if search_text:
                    haystack = " ".join(str(value or "") for value in (value, formula_text))
                    comparable = haystack if match_case else haystack.lower()
                    if search_needle not in comparable:
                        continue
                result = {
                    "address": address,
                    "value": value,
                    "formula": formula.text if formula is not None else None,
                    "styleIndex": int(cell.attrib.get("s", "0")),
                    "type": cell.attrib.get("t"),
                }
                if include_dependencies and formula_text:
                    result["dependencies"] = excel_formula_dependencies(formula_text, sheet_name)
                    dependency_edges.extend({"from": {"sheet": sheet_name, "address": address}, "to": dependency} for dependency in result["dependencies"])
                if include_cells:
                    cells.append(result)
                if search_text and len(matches) < max_matches:
                    matches.append({"sheet": sheet_name, **result})
                total_cells += 1
                if search_text and len(matches) >= max_matches:
                    break
        sheet_result["cells"] = cells
        sheet_result["cellCount"] = len(cells)
        sheet_result["tables"] = tables
        sheet_result["tableCount"] = len(tables)
        sheet_result["charts"] = charts
        sheet_result["chartCount"] = len(charts)
        sheet_result["pivots"] = pivots
        sheet_result["pivotCount"] = len(pivots)
        sheets.append(sheet_result)
        if search_text and len(matches) >= max_matches:
            break
    missing_sheets = sorted(requested_sheets - found_sheet_names)
    if missing_sheets:
        raise OfficePackageError("Excel worksheet was not found: %s" % ", ".join(missing_sheets))
    defined_names = []
    names_node = workbook.find(q("s", "definedNames"))
    for name in list(names_node or []):
        defined_names.append({"name": name.attrib.get("name"), "value": name.text})
    return {
        "kind": "excel",
        "sheets": sheets,
        "sheetCount": len(sheets),
        "cellCount": total_cells,
        "tableCount": total_tables,
        "chartCount": total_charts,
        "pivotCount": total_pivots,
        "formulaDependencies": dependency_edges,
        "formulaDependencyCount": len(dependency_edges),
        "definedNames": defined_names,
        "selectors": {
            "sheetNames": sorted(requested_sheets),
            "address": requested_range or None,
            "searchText": search_text or None,
            "matchCase": match_case,
        },
        "matches": matches,
        "matchCount": len(matches),
        "truncated": total_cells >= max_cells or (bool(search_text) and len(matches) >= max_matches),
    }


def ppt_shape_result(shape: ET.Element, slide_index: int, relationships: Optional[Dict[str, Dict[str, str]]] = None) -> Dict[str, Any]:
    nv = shape.find(".//" + q("p", "cNvPr"))
    placeholder = shape.find(".//" + q("p", "ph"))
    transform = shape.find(q("p", "xfrm"))
    if transform is None:
        transform = shape.find(".//" + q("a", "xfrm"))
    offset = transform.find(q("a", "off")) if transform is not None else None
    extent = transform.find(q("a", "ext")) if transform is not None else None
    paragraphs = []
    for paragraph in shape.iter(q("a", "p")):
        runs = []
        for run in list(paragraph):
            if run.tag in {q("a", "r"), q("a", "fld")}:
                props = run.find(q("a", "rPr"))
                runs.append({"text": text_of(run, (q("a", "t"),)), "properties": dict(props.attrib) if props is not None else {}})
        paragraphs.append({"text": text_of(paragraph, (q("a", "t"),)), "runs": runs})
    table = shape.find(".//" + q("a", "tbl"))
    table_rows = []
    if table is not None:
        for row in table.findall(q("a", "tr")):
            table_rows.append([text_of(cell, (q("a", "t"),)) for cell in row.findall(q("a", "tc"))])
    blip = shape.find(".//" + q("a", "blip"))
    image_rel_id = blip.attrib.get(q("r", "embed")) if blip is not None else None
    image_rel = (relationships or {}).get(image_rel_id or "")
    result = {
        "slideIndex": slide_index,
        "type": shape.tag.split("}")[-1],
        "id": nv.attrib.get("id") if nv is not None else None,
        "name": nv.attrib.get("name") if nv is not None else None,
        "description": nv.attrib.get("descr") if nv is not None else None,
        "text": text_of(shape, (q("a", "t"),)),
        "placeholder": dict(placeholder.attrib) if placeholder is not None else None,
        "geometry": {
            "x": int(offset.attrib.get("x", "0")) if offset is not None else None,
            "y": int(offset.attrib.get("y", "0")) if offset is not None else None,
            "width": int(extent.attrib.get("cx", "0")) if extent is not None else None,
            "height": int(extent.attrib.get("cy", "0")) if extent is not None else None,
            "rotation": int(transform.attrib.get("rot", "0")) if transform is not None else 0,
            "flipHorizontal": transform.attrib.get("flipH") in {"1", "true"} if transform is not None else False,
            "flipVertical": transform.attrib.get("flipV") in {"1", "true"} if transform is not None else False,
        },
        "paragraphs": paragraphs,
        "hasTable": table is not None,
        "table": {"rows": table_rows, "rowCount": len(table_rows), "columnCount": max([len(row) for row in table_rows] or [0])} if table is not None else None,
        "hasImage": shape.tag == q("p", "pic"),
        "image": {"relationshipId": image_rel_id, "target": image_rel.get("Target") if image_rel else None, "external": image_rel.get("TargetMode") == "External" if image_rel else False} if image_rel_id else None,
    }
    if shape.tag == q("p", "grpSp"):
        result["children"] = [ppt_shape_result(child, slide_index, relationships) for child in list(shape) if child.tag in {q("p", "sp"), q("p", "pic"), q("p", "graphicFrame"), q("p", "grpSp")}]
    return result


def inspect_powerpoint(package: zipfile.ZipFile, request: Dict[str, Any]) -> Dict[str, Any]:
    presentation = xml_root(package, "ppt/presentation.xml")
    relationships = rels_for_part(package, "ppt/presentation.xml")
    slide_list = presentation.find(q("p", "sldIdLst"))
    max_slides = max(1, min(int(request.get("maxSlides", 500)), 5000))
    slides = []
    for slide_node in list(slide_list or [])[:max_slides]:
        rel_id = slide_node.attrib.get(q("r", "id"))
        relationship = relationships.get(rel_id or "", {})
        target = relationship.get("Target", "")
        part = resolved_relationship_target("ppt/presentation.xml", target) if target else None
        slide_result: Dict[str, Any] = {
            "index": len(slides),
            "slideId": slide_node.attrib.get("id"),
            "relationshipId": rel_id,
            "part": part,
            "shapes": [],
            "notes": None,
        }
        if part and part in package.namelist():
            root = xml_root(package, part)
            slide_relationships = rels_for_part(package, part)
            shape_tree = root.find(".//" + q("p", "spTree"))
            if shape_tree is not None:
                for child in list(shape_tree):
                    if child.tag in {q("p", "sp"), q("p", "pic"), q("p", "graphicFrame"), q("p", "grpSp")}:
                        slide_result["shapes"].append(ppt_shape_result(child, len(slides), slide_relationships))
            for rel in slide_relationships.values():
                if rel.get("Type", "").endswith("/notesSlide"):
                    notes_part = resolved_relationship_target(part, rel.get("Target", ""))
                    if notes_part in package.namelist():
                        slide_result["notes"] = text_of(xml_root(package, notes_part), (q("a", "t"),))
                    break
        slide_result["shapeCount"] = len(slide_result["shapes"])
        slides.append(slide_result)
    size = presentation.find(q("p", "sldSz"))
    return {
        "kind": "powerpoint",
        "slides": slides,
        "slideCount": len(slides),
        "slideSize": {
            "width": int(size.attrib.get("cx", "0")) if size is not None else None,
            "height": int(size.attrib.get("cy", "0")) if size is not None else None,
            "type": size.attrib.get("type") if size is not None else None,
        },
        "truncated": len(list(slide_list or [])) > max_slides,
    }


def inspect_package(path: Path, request: Dict[str, Any]) -> Dict[str, Any]:
    if not path.is_file():
        raise OfficePackageError("Office package does not exist: %s" % path)
    try:
        with zipfile.ZipFile(path, "r") as package:
            inventory = package_inventory(package)
            kind = detect_kind(package, inventory["names"])
            requested_kind = request.get("kind")
            if requested_kind and requested_kind != kind:
                raise OfficePackageError("Expected a %s package but detected %s." % (requested_kind, kind))
            relationship_errors = validate_relationships(package, inventory["names"])
            if relationship_errors and request.get("strictRelationships", True):
                raise OfficePackageError("Office package has broken internal relationships.")
            content_type_errors = validate_content_types(package, inventory["names"])
            if content_type_errors and request.get("strictContentTypes", True):
                raise OfficePackageError("Office package has missing or incomplete content types.")
            if kind == "word":
                content = inspect_word(package, request)
            elif kind == "excel":
                content = inspect_excel(package, request)
            else:
                content = inspect_powerpoint(package, request)
            search_text = request.get("searchText")
            if isinstance(search_text, str) and search_text:
                needle = search_text if request.get("matchCase", False) else search_text.lower()
                max_matches = max(1, min(int(request.get("maxMatches", 200)), 5000))
                matches = []
                if kind == "word":
                    candidates = [
                        *({"objectType": "paragraph", **entry} for entry in content.get("paragraphs", [])),
                        *({"objectType": "contentControl", **entry} for entry in content.get("contentControls", [])),
                        *({"objectType": "comment", **entry} for entry in content.get("comments", [])),
                    ]
                elif kind == "excel":
                    candidates = [
                        {"objectType": "cell", "sheet": sheet.get("name"), **cell}
                        for sheet in content.get("sheets", []) for cell in sheet.get("cells", [])
                    ]
                else:
                    candidates = []
                    for slide in content.get("slides", []):
                        candidates.extend({"objectType": "shape", "slideIndex": slide.get("index"), **shape} for shape in slide.get("shapes", []))
                        if slide.get("notes"):
                            candidates.append({"objectType": "notes", "slideIndex": slide.get("index"), "text": slide.get("notes")})
                for candidate in candidates:
                    text_value = str(candidate.get("text", candidate.get("value", "")) or "")
                    haystack = text_value if request.get("matchCase", False) else text_value.lower()
                    if needle in haystack:
                        matches.append(candidate)
                        if len(matches) >= max_matches:
                            break
                content["search"] = {"query": search_text, "matches": matches, "matchCount": len(matches), "truncated": len(matches) >= max_matches}
            content.update({
                "package": {
                    "path": str(path),
                    "bytes": path.stat().st_size,
                    "entryCount": inventory["entryCount"],
                    "uncompressedBytes": inventory["uncompressedBytes"],
                    "hasMacros": inventory["hasMacros"],
                    "hasDigitalSignatures": inventory["hasDigitalSignatures"],
                    "relationshipErrors": relationship_errors,
                    "contentTypeErrors": content_type_errors,
                    "fingerprint": package_fingerprint(package),
                }
            })
            return content
    except zipfile.BadZipFile as error:
        raise OfficePackageError("File is not a readable Open XML package; it may be encrypted or a legacy Office format.") from error


def validate_package(path: Path, request: Dict[str, Any]) -> Dict[str, Any]:
    result = inspect_package(path, {**request, "includeCells": False, "maxParagraphs": 1, "maxSlides": 1})
    return {
        "valid": True,
        "kind": result["kind"],
        "package": result["package"],
    }


def replace_text_nodes(nodes: List[ET.Element], find: str, replacement: str, match_case: bool, replace_all: bool) -> List[Dict[str, Any]]:
    if not find:
        raise OfficePackageError("replaceText.find must not be empty.")
    if find == replacement:
        return []
    changes = []
    search_from = 0
    max_replacements = 100_000
    while True:
        combined = "".join(node.text or "" for node in nodes)
        haystack = combined if match_case else combined.lower()
        needle = find if match_case else find.lower()
        start = haystack.find(needle, search_from)
        if start < 0:
            break
        end = start + len(find)
        positions = []
        cursor = 0
        for index, node in enumerate(nodes):
            length = len(node.text or "")
            positions.append((index, cursor, cursor + length))
            cursor += length
        start_entry = next((entry for entry in positions if entry[1] <= start < entry[2]), None)
        end_entry = next((entry for entry in positions if entry[1] < end <= entry[2]), None)
        if start_entry is None or end_entry is None:
            break
        start_index, start_base, _ = start_entry
        end_index, end_base, _ = end_entry
        start_text = nodes[start_index].text or ""
        end_text = nodes[end_index].text or ""
        if start_index == end_index:
            nodes[start_index].text = start_text[: start - start_base] + replacement + start_text[end - end_base :]
        else:
            nodes[start_index].text = start_text[: start - start_base] + replacement
            for index in range(start_index + 1, end_index):
                nodes[index].text = ""
            nodes[end_index].text = end_text[end - end_base :]
        changes.append({"before": combined[start:end], "after": replacement, "offset": start})
        if len(changes) >= max_replacements:
            raise OfficePackageError("Text replacement exceeded the 100,000-match safety limit.")
        if not replace_all:
            break
        search_from = start + len(replacement)
    return changes


def set_text_nodes(nodes: List[ET.Element], value: Any) -> str:
    """Replace visible text while retaining the first run's formatting."""
    before = "".join(node.text or "" for node in nodes)
    if not nodes:
        raise OfficePackageError("The selected Office object has no editable text runs.")
    nodes[0].text = "" if value is None else str(value)
    for node in nodes[1:]:
        node.text = ""
    return before


def word_part(package: zipfile.ZipFile, modifications: Dict[str, bytes], part: str) -> ET.Element:
    if part not in package.namelist():
        raise OfficePackageError("Word package part was not found: %s" % part)
    return ET.fromstring(modifications.get(part, package.read(part)))


def word_paragraphs(root: ET.Element) -> List[ET.Element]:
    return list(root.iter(q("w", "p")))


def new_word_paragraph(text: Any, style: Optional[str] = None) -> ET.Element:
    paragraph = ET.Element(q("w", "p"))
    if style:
        props = ET.SubElement(paragraph, q("w", "pPr"))
        ET.SubElement(props, q("w", "pStyle"), {q("w", "val"): str(style)})
    run = ET.SubElement(paragraph, q("w", "r"))
    ET.SubElement(run, q("w", "t")).text = "" if text is None else str(text)
    return paragraph


def direct_parent(root: ET.Element, child: ET.Element) -> Optional[ET.Element]:
    return next((parent for parent in root.iter() if child in list(parent)), None)


def word_relationship_part(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def word_relationship_root(package: zipfile.ZipFile, modifications: Dict[str, bytes], part: str) -> Tuple[str, ET.Element]:
    rel_part = word_relationship_part(part)
    if rel_part in modifications:
        return rel_part, ET.fromstring(modifications[rel_part])
    if rel_part in package.namelist():
        return rel_part, ET.fromstring(package.read(rel_part))
    return rel_part, ET.Element(q("pr", "Relationships"))


def add_word_relationship(package: zipfile.ZipFile, modifications: Dict[str, bytes], part: str, rel_type: str, target: str, external: bool = False) -> str:
    rel_part, root = word_relationship_root(package, modifications, part)
    for relationship in root.findall(q("pr", "Relationship")):
        if relationship.attrib.get("Type") == rel_type and relationship.attrib.get("Target") == target and (relationship.attrib.get("TargetMode") == "External") == external:
            return relationship.attrib["Id"]
    used_ids = {relationship.attrib.get("Id") for relationship in root.findall(q("pr", "Relationship"))}
    number = 1
    while "rId%d" % number in used_ids:
        number += 1
    rel_id = "rId%d" % number
    attributes = {"Id": rel_id, "Type": rel_type, "Target": target}
    if external:
        attributes["TargetMode"] = "External"
    ET.SubElement(root, q("pr", "Relationship"), attributes)
    modifications[rel_part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return rel_id


def add_content_type_override(package: zipfile.ZipFile, modifications: Dict[str, bytes], part_name: str, content_type: str) -> None:
    part = "[Content_Types].xml"
    root = ET.fromstring(modifications.get(part, package.read(part)))
    existing = next((entry for entry in root.findall(q("ct", "Override")) if entry.attrib.get("PartName") == part_name), None)
    if existing is None:
        ET.SubElement(root, q("ct", "Override"), {"PartName": part_name, "ContentType": content_type})
    else:
        existing.attrib["ContentType"] = content_type
    modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def assert_word_has_no_tracked_changes(package: zipfile.ZipFile, modifications: Dict[str, bytes]) -> None:
    tracked_tags = {q("w", name) for name in ("ins", "del", "moveFrom", "moveTo")}
    parts = sorted(name for name in package.namelist() if re.fullmatch(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml", name))
    for part in parts:
        root = ET.fromstring(modifications.get(part, package.read(part)))
        if any(node.tag in tracked_tags for node in root.iter()):
            raise OfficePackageError("Word documents containing tracked changes are refused. Accept or reject tracked changes in Word before editing.")


def new_word_table(rows: List[List[str]], style: Optional[str] = None) -> ET.Element:
    if not rows or any(not row for row in rows):
        raise OfficePackageError("Word insertTable requires at least one non-empty row.")
    if len(rows) > 100 or max(len(row) for row in rows) > 50 or sum(len(row) for row in rows) > 5000:
        raise OfficePackageError("Word insertTable exceeds the 100-row, 50-column, or 5,000-cell safety limit.")
    width = max(len(row) for row in rows)
    table = ET.Element(q("w", "tbl"))
    if style:
        props = ET.SubElement(table, q("w", "tblPr"))
        ET.SubElement(props, q("w", "tblStyle"), {q("w", "val"): str(style)})
    grid = ET.SubElement(table, q("w", "tblGrid"))
    for _ in range(width):
        ET.SubElement(grid, q("w", "gridCol"))
    for values in rows:
        row = ET.SubElement(table, q("w", "tr"))
        for value in list(values) + [""] * (width - len(values)):
            cell = ET.SubElement(row, q("w", "tc"))
            cell.append(new_word_paragraph(value))
    return table


def edit_word(package: zipfile.ZipFile, request: Dict[str, Any]) -> Tuple[Dict[str, bytes], List[Dict[str, Any]]]:
    names = set(package.namelist())
    modifications: Dict[str, bytes] = {}
    changes: List[Dict[str, Any]] = []
    if request.get("trackedChanges", "refuse") != "refuse":
        raise OfficePackageError("trackedChanges must be refuse; preserving review semantics is not supported.")
    assert_word_has_no_tracked_changes(package, modifications)
    for operation in request.get("operations", []):
        op_type = operation.get("type")
        if op_type == "addHyperlink":
            part = str(operation.get("part", "word/document.xml"))
            root = word_part(package, modifications, part)
            paragraphs = word_paragraphs(root)
            index = int(operation.get("paragraphIndex", -1))
            if index < 0 or index >= len(paragraphs):
                raise OfficePackageError("Word paragraphIndex is out of range.")
            target = str(operation.get("url", "")).strip()
            parsed = urlparse(target)
            if parsed.scheme.lower() not in {"http", "https", "mailto"} or (parsed.scheme.lower() in {"http", "https"} and not parsed.netloc):
                raise OfficePackageError("Word addHyperlink url must use http, https, or mailto.")
            rel_id = add_word_relationship(package, modifications, part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target, external=True)
            hyperlink = ET.SubElement(paragraphs[index], q("w", "hyperlink"), {q("r", "id"): rel_id})
            run = ET.SubElement(hyperlink, q("w", "r"))
            run_props = ET.SubElement(run, q("w", "rPr"))
            ET.SubElement(run_props, q("w", "rStyle"), {q("w", "val"): "Hyperlink"})
            ET.SubElement(run, q("w", "t")).text = str(operation.get("text", ""))
            changes.append({"operation": op_type, "part": part, "paragraphIndex": index, "text": operation.get("text", ""), "url": target})
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type == "addComment":
            part = "word/document.xml"
            root = word_part(package, modifications, part)
            paragraphs = word_paragraphs(root)
            index = int(operation.get("paragraphIndex", -1))
            if index < 0 or index >= len(paragraphs):
                raise OfficePackageError("Word paragraphIndex is out of range.")
            comments_part = "word/comments.xml"
            if comments_part in modifications:
                comments_root = ET.fromstring(modifications[comments_part])
            elif comments_part in names:
                comments_root = ET.fromstring(package.read(comments_part))
            else:
                comments_root = ET.Element(q("w", "comments"))
            used_ids = {int(node.attrib.get(q("w", "id"), "-1")) for node in comments_root.findall(q("w", "comment")) if node.attrib.get(q("w", "id"), "").isdigit()}
            comment_id = max(used_ids, default=-1) + 1
            author = str(operation.get("author", "Codex"))
            attributes = {q("w", "id"): str(comment_id), q("w", "author"): author, q("w", "date"): datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")}
            if operation.get("initials") is not None:
                attributes[q("w", "initials")] = str(operation.get("initials"))
            comment = ET.SubElement(comments_root, q("w", "comment"), attributes)
            comment.append(new_word_paragraph(operation.get("text")))
            modifications[comments_part] = ET.tostring(comments_root, encoding="utf-8", xml_declaration=True)
            add_word_relationship(package, modifications, part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml")
            add_content_type_override(package, modifications, "/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
            paragraph = paragraphs[index]
            start_index = 1 if len(paragraph) and paragraph[0].tag == q("w", "pPr") else 0
            paragraph.insert(start_index, ET.Element(q("w", "commentRangeStart"), {q("w", "id"): str(comment_id)}))
            paragraph.append(ET.Element(q("w", "commentRangeEnd"), {q("w", "id"): str(comment_id)}))
            reference_run = ET.SubElement(paragraph, q("w", "r"))
            ET.SubElement(reference_run, q("w", "commentReference"), {q("w", "id"): str(comment_id)})
            changes.append({"operation": op_type, "part": part, "paragraphIndex": index, "commentId": str(comment_id), "author": author})
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type == "insertTable":
            part = "word/document.xml"
            root = word_part(package, modifications, part)
            body = root.find(q("w", "body"))
            if body is None:
                raise OfficePackageError("Word document body is missing.")
            table = new_word_table(operation.get("rows", []), operation.get("style"))
            after = operation.get("afterParagraphIndex")
            if after is None:
                section = body.find(q("w", "sectPr"))
                body.insert(list(body).index(section) if section is not None else len(body), table)
            else:
                paragraphs = word_paragraphs(root)
                index = int(after)
                if index < 0 or index >= len(paragraphs) or direct_parent(root, paragraphs[index]) is not body:
                    raise OfficePackageError("Word afterParagraphIndex must identify a top-level document paragraph.")
                body.insert(list(body).index(paragraphs[index]) + 1, table)
            changes.append({"operation": op_type, "part": part, "afterParagraphIndex": after, "rowCount": len(operation.get("rows", [])), "columnCount": max(len(row) for row in operation.get("rows", []))})
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type in {"setParagraphText", "setParagraphStyle", "insertParagraph"}:
            part = str(operation.get("part", "word/document.xml"))
            root = word_part(package, modifications, part)
            paragraphs = word_paragraphs(root)
            if op_type == "insertParagraph":
                after = operation.get("afterIndex")
                if after is None:
                    body = root.find(q("w", "body"))
                    if body is None:
                        raise OfficePackageError("Word document body is missing.")
                    paragraph = new_word_paragraph(operation.get("text"), operation.get("style"))
                    section = body.find(q("w", "sectPr"))
                    body.insert(list(body).index(section) if section is not None else len(body), paragraph)
                    index = len(paragraphs)
                else:
                    index = int(after)
                    if index < 0 or index >= len(paragraphs):
                        raise OfficePackageError("Word afterIndex is out of range.")
                    anchor = paragraphs[index]
                    parent = direct_parent(root, anchor)
                    if parent is None:
                        raise OfficePackageError("Could not locate the Word paragraph parent.")
                    paragraph = new_word_paragraph(operation.get("text"), operation.get("style"))
                    parent.insert(list(parent).index(anchor) + 1, paragraph)
                    index += 1
                changes.append({"operation": op_type, "part": part, "paragraphIndex": index, "after": operation.get("text", "")})
            else:
                index = int(operation.get("paragraphIndex", -1))
                if index < 0 or index >= len(paragraphs):
                    raise OfficePackageError("Word paragraphIndex is out of range.")
                paragraph = paragraphs[index]
                if op_type == "setParagraphText":
                    before = set_text_nodes(list(paragraph.iter(q("w", "t"))), operation.get("text"))
                    changes.append({"operation": op_type, "part": part, "paragraphIndex": index, "before": before, "after": operation.get("text", "")})
                else:
                    ppr = paragraph.find(q("w", "pPr"))
                    if ppr is None:
                        ppr = ET.Element(q("w", "pPr"))
                        paragraph.insert(0, ppr)
                    style = ppr.find(q("w", "pStyle"))
                    before = style.attrib.get(q("w", "val")) if style is not None else None
                    if style is None:
                        style = ET.SubElement(ppr, q("w", "pStyle"))
                    style.attrib[q("w", "val")] = str(operation.get("style", "Normal"))
                    changes.append({"operation": op_type, "part": part, "paragraphIndex": index, "before": before, "after": operation.get("style", "Normal")})
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type == "setTableCell":
            part = str(operation.get("part", "word/document.xml"))
            root = word_part(package, modifications, part)
            tables = list(root.iter(q("w", "tbl")))
            table_index = int(operation.get("tableIndex", -1))
            row_index = int(operation.get("rowIndex", -1))
            column_index = int(operation.get("columnIndex", -1))
            if table_index < 0 or table_index >= len(tables):
                raise OfficePackageError("Word tableIndex is out of range.")
            rows = tables[table_index].findall(q("w", "tr"))
            if row_index < 0 or row_index >= len(rows):
                raise OfficePackageError("Word rowIndex is out of range.")
            cells = rows[row_index].findall(q("w", "tc"))
            if column_index < 0 or column_index >= len(cells):
                raise OfficePackageError("Word columnIndex is out of range.")
            nodes = list(cells[column_index].iter(q("w", "t")))
            if not nodes:
                paragraph = cells[column_index].find(q("w", "p")) or ET.SubElement(cells[column_index], q("w", "p"))
                run = ET.SubElement(paragraph, q("w", "r"))
                nodes = [ET.SubElement(run, q("w", "t"))]
            before = set_text_nodes(nodes, operation.get("text"))
            changes.append({"operation": op_type, "part": part, "tableIndex": table_index, "rowIndex": row_index, "columnIndex": column_index, "before": before, "after": operation.get("text", "")})
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type == "setContentControlText":
            part = str(operation.get("part", "word/document.xml"))
            root = word_part(package, modifications, part)
            controls = list(root.iter(q("w", "sdt")))
            requested_index = operation.get("contentControlIndex")
            selected = None
            for index, control in enumerate(controls):
                props = control.find(q("w", "sdtPr"))
                tag_node = props.find(q("w", "tag")) if props is not None else None
                id_node = props.find(q("w", "id")) if props is not None else None
                if requested_index is not None and index != int(requested_index):
                    continue
                if operation.get("tag") is not None and (tag_node is None or tag_node.attrib.get(q("w", "val")) != str(operation.get("tag"))):
                    continue
                if operation.get("id") is not None and (id_node is None or id_node.attrib.get(q("w", "val")) != str(operation.get("id"))):
                    continue
                selected = (index, control)
                break
            if selected is None:
                raise OfficePackageError("Word content control was not found.")
            index, control = selected
            before = set_text_nodes(list(control.iter(q("w", "t"))), operation.get("text"))
            changes.append({"operation": op_type, "part": part, "contentControlIndex": index, "before": before, "after": operation.get("text", "")})
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type != "replaceText":
            raise OfficePackageError("Unsupported Word operation: %s" % op_type)
        scope = operation.get("scope", "document")
        parts = ["word/document.xml"] if scope in {"document", "all"} else []
        if scope in {"headers", "all"}:
            parts += sorted(name for name in names if re.fullmatch(r"word/header\d+\.xml", name))
        if scope in {"footers", "all"}:
            parts += sorted(name for name in names if re.fullmatch(r"word/footer\d+\.xml", name))
        if scope not in {"document", "headers", "footers", "all"}:
            raise OfficePackageError("Word replaceText.scope is invalid.")
        operation_finished = False
        for part in dict.fromkeys(parts):
            root = ET.fromstring(modifications.get(part, package.read(part)))
            for paragraph_index, paragraph in enumerate(root.iter(q("w", "p"))):
                text_nodes = list(paragraph.iter(q("w", "t")))
                part_changes = replace_text_nodes(
                    text_nodes,
                    str(operation.get("find", "")),
                    str(operation.get("replace", "")),
                    bool(operation.get("matchCase", True)),
                    bool(operation.get("all", True)),
                )
                for change in part_changes:
                    changes.append({"operation": "replaceText", "part": part, "paragraphIndex": paragraph_index, **change})
                if part_changes and not operation.get("all", True):
                    operation_finished = True
                    break
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            if operation_finished:
                break
    return modifications, changes


def excel_sheet_parts(package: zipfile.ZipFile) -> Dict[str, str]:
    workbook = xml_root(package, "xl/workbook.xml")
    relationships = rels_for_part(package, "xl/workbook.xml")
    result = {}
    for sheet in list(workbook.find(q("s", "sheets")) or []):
        rel = relationships.get(sheet.attrib.get(q("r", "id"), ""), {})
        if rel.get("Target"):
            result[sheet.attrib.get("name", "")] = resolved_relationship_target("xl/workbook.xml", rel["Target"])
    return result


def excel_row_number(address: str) -> int:
    match = re.fullmatch(r"[A-Za-z]{1,3}([1-9][0-9]*)", address)
    if not match:
        raise OfficePackageError("Invalid Excel cell address: %s" % address)
    return int(match.group(1))


def excel_address_parts(address: str) -> Tuple[int, int]:
    match = re.fullmatch(r"([A-Za-z]{1,3})([1-9][0-9]*)", address)
    if not match:
        raise OfficePackageError("Invalid Excel cell address: %s" % address)
    column = 0
    for character in match.group(1).upper():
        column = column * 26 + ord(character) - 64
    return column, int(match.group(2))


def excel_column_name(column: int) -> str:
    if column < 1 or column > 16384:
        raise OfficePackageError("Excel column is outside the worksheet limit.")
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def excel_range_addresses(address: str) -> List[List[str]]:
    parts = str(address).upper().split(":")
    if len(parts) == 1:
        excel_address_parts(parts[0])
        return [[parts[0]]]
    if len(parts) != 2:
        raise OfficePackageError("Invalid Excel range address: %s" % address)
    start_column, start_row = excel_address_parts(parts[0])
    end_column, end_row = excel_address_parts(parts[1])
    if end_column < start_column or end_row < start_row:
        raise OfficePackageError("Excel range end precedes its start.")
    if (end_column - start_column + 1) * (end_row - start_row + 1) > 100_000:
        raise OfficePackageError("Excel range exceeds the 100,000-cell edit limit.")
    return [
        ["%s%d" % (excel_column_name(column), row) for column in range(start_column, end_column + 1)]
        for row in range(start_row, end_row + 1)
    ]


def find_or_create_excel_cell(root: ET.Element, address: str) -> ET.Element:
    sheet_data = root.find(q("s", "sheetData"))
    if sheet_data is None:
        sheet_data = ET.SubElement(root, q("s", "sheetData"))
    row_number = excel_row_number(address)
    row = next((entry for entry in sheet_data.findall(q("s", "row")) if int(entry.attrib.get("r", "0")) == row_number), None)
    if row is None:
        row = ET.SubElement(sheet_data, q("s", "row"), {"r": str(row_number)})
    cell = next((entry for entry in row.findall(q("s", "c")) if entry.attrib.get("r") == address.upper()), None)
    if cell is None:
        cell = ET.SubElement(row, q("s", "c"), {"r": address.upper()})
    return cell


def excel_table_target(
    package: zipfile.ZipFile,
    modifications: Dict[str, Optional[bytes]],
    sheet_parts: Dict[str, str],
    table_name: str,
) -> Tuple[str, str, ET.Element, ET.Element]:
    matches = []
    for sheet_name, sheet_part in sheet_parts.items():
        for relationship in rels_for_part(package, sheet_part).values():
            if not relationship.get("Type", "").endswith("/table"):
                continue
            table_part = resolved_relationship_target(sheet_part, relationship.get("Target", ""))
            if table_part not in package.namelist():
                continue
            table = ET.fromstring(modifications.get(table_part) or package.read(table_part))
            if table_name in {table.attrib.get("name"), table.attrib.get("displayName")}:
                sheet = ET.fromstring(modifications.get(sheet_part) or package.read(sheet_part))
                matches.append((sheet_name, sheet_part, sheet, table_part, table))
    if not matches:
        raise OfficePackageError("Excel table was not found: %s" % table_name)
    if len(matches) > 1:
        raise OfficePackageError("Excel table name is ambiguous: %s" % table_name)
    _sheet_name, sheet_part, sheet, table_part, table = matches[0]
    return sheet_part, table_part, sheet, table


def excel_table_bounds(table: ET.Element) -> Tuple[int, int, int, int]:
    reference = str(table.attrib.get("ref", ""))
    endpoints = reference.split(":")
    if len(endpoints) != 2:
        raise OfficePackageError("Excel table has an invalid range: %s" % reference)
    start_column, start_row = excel_address_parts(endpoints[0])
    end_column, end_row = excel_address_parts(endpoints[1])
    if start_column > end_column or start_row > end_row:
        raise OfficePackageError("Excel table has an invalid range: %s" % reference)
    return start_column, start_row, end_column, end_row


def excel_cells_in_columns(root: ET.Element, row_number: int, start_column: int, end_column: int) -> List[ET.Element]:
    sheet_data = root.find(q("s", "sheetData"))
    row = next((entry for entry in list(sheet_data or []) if int(entry.attrib.get("r", "0")) == row_number), None)
    if row is None:
        return []
    result = []
    for cell in row.findall(q("s", "c")):
        address = cell.attrib.get("r", "")
        try:
            column, _row = excel_address_parts(address)
        except OfficePackageError:
            continue
        if start_column <= column <= end_column:
            result.append(cell)
    return result


def excel_remove_cell(root: ET.Element, address: str) -> None:
    sheet_data = root.find(q("s", "sheetData"))
    if sheet_data is None:
        return
    row_number = excel_row_number(address)
    row = next((entry for entry in sheet_data.findall(q("s", "row")) if int(entry.attrib.get("r", "0")) == row_number), None)
    if row is None:
        return
    cell = next((entry for entry in row.findall(q("s", "c")) if entry.attrib.get("r") == address.upper()), None)
    if cell is not None:
        row.remove(cell)
    if not row.findall(q("s", "c")):
        sheet_data.remove(row)


def excel_shift_table_cells_down(root: ET.Element, start_row: int, end_row: int, start_column: int, end_column: int, count: int) -> None:
    if end_row + count > 1_048_576:
        raise OfficePackageError("Excel table expansion exceeds the worksheet row limit.")
    for row_number in range(end_row + 1, end_row + count + 1):
        if excel_cells_in_columns(root, row_number, start_column, end_column):
            raise OfficePackageError("Excel table cannot expand because cells immediately below it are not empty.")
    for row_number in range(end_row, start_row - 1, -1):
        for column in range(start_column, end_column + 1):
            source_address = "%s%d" % (excel_column_name(column), row_number)
            source_cells = excel_cells_in_columns(root, row_number, column, column)
            if not source_cells:
                continue
            source = source_cells[0]
            destination_address = "%s%d" % (excel_column_name(column), row_number + count)
            destination = find_or_create_excel_cell(root, destination_address)
            destination.attrib.clear()
            destination.attrib.update(source.attrib)
            destination.attrib["r"] = destination_address
            destination[:] = [ET.fromstring(ET.tostring(child)) for child in list(source)]
            excel_remove_cell(root, source_address)


def excel_update_table_ref(table: ET.Element, start_column: int, start_row: int, end_column: int, end_row: int) -> None:
    reference = "%s%d:%s%d" % (excel_column_name(start_column), start_row, excel_column_name(end_column), end_row)
    table.attrib["ref"] = reference
    auto_filter = table.find(q("s", "autoFilter"))
    if auto_filter is not None:
        totals_count = int(table.attrib.get("totalsRowCount", "0"))
        filter_end = end_row - totals_count
        auto_filter.attrib["ref"] = "%s%d:%s%d" % (excel_column_name(start_column), start_row, excel_column_name(end_column), filter_end)


def excel_add_table_rows(
    package: zipfile.ZipFile,
    modifications: Dict[str, Optional[bytes]],
    sheet_parts: Dict[str, str],
    operation: Dict[str, Any],
) -> Dict[str, Any]:
    table_name = str(operation.get("table", ""))
    sheet_part, table_part, sheet, table = excel_table_target(package, modifications, sheet_parts, table_name)
    start_column, start_row, end_column, end_row = excel_table_bounds(table)
    column_count = end_column - start_column + 1
    values = operation.get("values")
    if not isinstance(values, list) or not values or any(not isinstance(row, list) or len(row) != column_count for row in values):
        raise OfficePackageError("Excel addTableRow values must contain exactly %d columns per row." % column_count)
    header_count = int(table.attrib.get("headerRowCount", "1"))
    totals_count = int(table.attrib.get("totalsRowCount", "0"))
    if header_count not in {0, 1} or totals_count not in {0, 1}:
        raise OfficePackageError("Excel table uses unsupported header or totals row counts.")
    data_count = end_row - start_row + 1 - header_count - totals_count
    requested_index = operation.get("index")
    index = data_count if requested_index is None else int(requested_index)
    if index < 0 or index > data_count:
        raise OfficePackageError("Excel addTableRow index is outside the table data rows.")
    insertion_row = start_row + header_count + index
    excel_shift_table_cells_down(sheet, insertion_row, end_row, start_column, end_column, len(values))
    for row_offset, row_values in enumerate(values):
        for column_offset, value in enumerate(row_values):
            address = "%s%d" % (excel_column_name(start_column + column_offset), insertion_row + row_offset)
            set_excel_cell(find_or_create_excel_cell(sheet, address), value)
    excel_update_table_ref(table, start_column, start_row, end_column, end_row + len(values))
    modifications[sheet_part] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
    modifications[table_part] = ET.tostring(table, encoding="utf-8", xml_declaration=True)
    return {"table": table_name, "sheetPart": sheet_part, "index": index, "rowCount": len(values), "beforeRef": "%s%d:%s%d" % (excel_column_name(start_column), start_row, excel_column_name(end_column), end_row), "afterRef": table.attrib["ref"], "preservedTotalsRow": totals_count == 1}


EXCEL_TOTAL_FUNCTIONS = {
    "average": 101,
    "countNums": 102,
    "count": 103,
    "max": 104,
    "min": 105,
    "stdDev": 107,
    "sum": 109,
    "var": 110,
}


def excel_set_table_totals(
    package: zipfile.ZipFile,
    modifications: Dict[str, Optional[bytes]],
    sheet_parts: Dict[str, str],
    operation: Dict[str, Any],
) -> Dict[str, Any]:
    table_name = str(operation.get("table", ""))
    sheet_part, table_part, sheet, table = excel_table_target(package, modifications, sheet_parts, table_name)
    start_column, start_row, end_column, end_row = excel_table_bounds(table)
    before_enabled = int(table.attrib.get("totalsRowCount", "0")) == 1
    enabled = bool(operation.get("enabled"))
    if enabled and not before_enabled:
        excel_shift_table_cells_down(sheet, end_row + 1, end_row, start_column, end_column, 1)
        end_row += 1
    elif not enabled and before_enabled:
        for column in range(start_column, end_column + 1):
            excel_remove_cell(sheet, "%s%d" % (excel_column_name(column), end_row))
        end_row -= 1
    table.attrib["totalsRowCount"] = "1" if enabled else "0"
    table.attrib["totalsRowShown"] = "1" if enabled else "0"
    columns_node = table.find(q("s", "tableColumns"))
    columns = list(columns_node or [])
    if len(columns) != end_column - start_column + 1:
        raise OfficePackageError("Excel table column metadata does not match its range.")
    configured = []
    seen_columns = set()
    for setting in operation.get("columns", []):
        selector = setting.get("column")
        if isinstance(selector, bool):
            raise OfficePackageError("Excel totals column must be a name or zero-based index.")
        if isinstance(selector, int):
            column_index = selector
        else:
            column_index = next((index for index, column in enumerate(columns) if column.attrib.get("name") == str(selector)), -1)
        if column_index < 0 or column_index >= len(columns):
            raise OfficePackageError("Excel totals column was not found: %s" % selector)
        if column_index in seen_columns:
            raise OfficePackageError("Excel totals column is configured more than once: %s" % selector)
        seen_columns.add(column_index)
        column = columns[column_index]
        for attribute in ("totalsRowFunction", "totalsRowLabel"):
            column.attrib.pop(attribute, None)
        function = setting.get("function")
        label = setting.get("label")
        formula = setting.get("formula")
        if label is not None and (function is not None or formula is not None):
            raise OfficePackageError("Excel totals column cannot combine label with function or formula.")
        if formula is not None and function != "custom":
            raise OfficePackageError("Excel totals formula requires function custom.")
        if function == "custom" and not formula:
            raise OfficePackageError("Excel custom totals require formula.")
        if label is not None:
            column.attrib["totalsRowLabel"] = str(label)
        elif function is not None:
            column.attrib["totalsRowFunction"] = str(function)
        if enabled:
            cell = find_or_create_excel_cell(sheet, "%s%d" % (excel_column_name(start_column + column_index), end_row))
            if label is not None:
                set_excel_cell(cell, label)
            elif function == "custom":
                set_excel_cell(cell, formula=str(formula).lstrip("="))
            elif function in EXCEL_TOTAL_FUNCTIONS:
                escaped_name = str(column.attrib.get("name", "")).replace("]", "]]" )
                set_excel_cell(cell, formula="SUBTOTAL(%d,[%s])" % (EXCEL_TOTAL_FUNCTIONS[function], escaped_name))
            else:
                set_excel_cell(cell, None)
        configured.append({"column": selector, "function": function, "label": label})
    if enabled:
        configured_indexes = {
            setting.get("column") if isinstance(setting.get("column"), int) else next((index for index, column in enumerate(columns) if column.attrib.get("name") == str(setting.get("column"))), -1)
            for setting in operation.get("columns", [])
        }
        for index, column in enumerate(columns):
            if index not in configured_indexes and not before_enabled:
                column.attrib.pop("totalsRowFunction", None)
                column.attrib.pop("totalsRowLabel", None)
                set_excel_cell(find_or_create_excel_cell(sheet, "%s%d" % (excel_column_name(start_column + index), end_row)), None)
    excel_update_table_ref(table, start_column, start_row, end_column, end_row)
    modifications[sheet_part] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
    modifications[table_part] = ET.tostring(table, encoding="utf-8", xml_declaration=True)
    return {"table": table_name, "beforeEnabled": before_enabled, "afterEnabled": enabled, "afterRef": table.attrib["ref"], "columns": configured}


def set_excel_cell(cell: ET.Element, value: Any = None, formula: Optional[str] = None) -> None:
    for child in list(cell):
        if child.tag in {q("s", "v"), q("s", "f"), q("s", "is")}:
            cell.remove(child)
    if formula is not None:
        cell.attrib.pop("t", None)
        ET.SubElement(cell, q("s", "f")).text = formula
        if value is not None:
            ET.SubElement(cell, q("s", "v")).text = str(value)
    elif isinstance(value, bool):
        cell.attrib["t"] = "b"
        ET.SubElement(cell, q("s", "v")).text = "1" if value else "0"
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, q("s", "v")).text = str(value)
    else:
        cell.attrib["t"] = "inlineStr"
        inline = ET.SubElement(cell, q("s", "is"))
        ET.SubElement(inline, q("s", "t")).text = "" if value is None else str(value)


def excel_style_with_number_format(
    package: zipfile.ZipFile,
    modifications: Dict[str, Optional[bytes]],
    style_index: int,
    format_code: str,
) -> int:
    styles_part = "xl/styles.xml"
    if styles_part not in package.namelist():
        raise OfficePackageError("Excel setNumberFormat requires an existing styles.xml part.")
    if not format_code or len(format_code) > 255:
        raise OfficePackageError("Excel formatCode must contain 1 to 255 characters.")
    root = ET.fromstring(modifications.get(styles_part) or package.read(styles_part))
    number_formats = root.find(q("s", "numFmts"))
    if number_formats is None:
        number_formats = ET.Element(q("s", "numFmts"), {"count": "0"})
        root.insert(0, number_formats)
    existing_format = next((entry for entry in list(number_formats) if entry.attrib.get("formatCode") == format_code), None)
    if existing_format is None:
        used_ids = {int(entry.attrib.get("numFmtId", "0")) for entry in list(number_formats)}
        number_format_id = 164
        while number_format_id in used_ids:
            number_format_id += 1
        existing_format = ET.SubElement(number_formats, q("s", "numFmt"), {
            "numFmtId": str(number_format_id),
            "formatCode": format_code,
        })
        number_formats.attrib["count"] = str(len(list(number_formats)))
    number_format_id = existing_format.attrib["numFmtId"]
    cell_formats = root.find(q("s", "cellXfs"))
    formats = list(cell_formats or [])
    if cell_formats is None or style_index < 0 or style_index >= len(formats):
        raise OfficePackageError("Excel cell styleIndex is outside styles.xml cellXfs.")
    desired_attributes = {**formats[style_index].attrib, "numFmtId": number_format_id, "applyNumberFormat": "1"}
    existing_index = next((index for index, entry in enumerate(formats) if entry.attrib == desired_attributes), None)
    if existing_index is not None:
        return existing_index
    new_format = ET.fromstring(ET.tostring(formats[style_index]))
    new_format.attrib.clear()
    new_format.attrib.update(desired_attributes)
    cell_formats.append(new_format)
    cell_formats.attrib["count"] = str(len(list(cell_formats)))
    modifications[styles_part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return len(formats)


def apply_excel_recalculate(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]]) -> Dict[str, Any]:
    workbook_part = "xl/workbook.xml"
    workbook = ET.fromstring(modifications.get(workbook_part) or package.read(workbook_part))
    calculation = workbook.find(q("s", "calcPr"))
    before = dict(calculation.attrib) if calculation is not None else None
    if calculation is None:
        calculation = ET.SubElement(workbook, q("s", "calcPr"))
    calculation.attrib.update({"calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"})
    modifications[workbook_part] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)

    removed_chain = "xl/calcChain.xml" in package.namelist()
    if removed_chain:
        modifications["xl/calcChain.xml"] = None
    relationships_part = "xl/_rels/workbook.xml.rels"
    if relationships_part in package.namelist():
        relationships = ET.fromstring(modifications.get(relationships_part) or package.read(relationships_part))
        removed_relationship = False
        for relationship in list(relationships):
            target = relationship.attrib.get("Target", "")
            if relationship.attrib.get("Type", "").endswith("/calcChain") or resolved_relationship_target(workbook_part, target) == "xl/calcChain.xml":
                relationships.remove(relationship)
                removed_relationship = True
        if removed_relationship:
            modifications[relationships_part] = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
    if "[Content_Types].xml" in package.namelist():
        content_types = ET.fromstring(modifications.get("[Content_Types].xml") or package.read("[Content_Types].xml"))
        removed_override = False
        for entry in list(content_types):
            if entry.attrib.get("PartName") == "/xl/calcChain.xml":
                content_types.remove(entry)
                removed_override = True
        if removed_override:
            modifications["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    return {"before": before, "after": dict(calculation.attrib), "removedCalcChain": removed_chain}


def excel_add_fill_dxf(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], color: str) -> int:
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
        raise OfficePackageError("Excel conditional-format color must be six hexadecimal characters.")
    styles_part = "xl/styles.xml"
    if styles_part not in package.namelist():
        raise OfficePackageError("Excel conditional formatting requires an existing styles.xml part.")
    root = ET.fromstring(modifications.get(styles_part) or package.read(styles_part))
    differential = root.find(q("s", "dxfs"))
    if differential is None:
        differential = ET.SubElement(root, q("s", "dxfs"), {"count": "0"})
    dxf = ET.SubElement(differential, q("s", "dxf"))
    fill = ET.SubElement(dxf, q("s", "fill"))
    pattern = ET.SubElement(fill, q("s", "patternFill"), {"patternType": "solid"})
    ET.SubElement(pattern, q("s", "fgColor"), {"rgb": "FF" + color.upper()})
    ET.SubElement(pattern, q("s", "bgColor"), {"indexed": "64"})
    differential.attrib["count"] = str(len(list(differential)))
    modifications[styles_part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return len(list(differential)) - 1


def excel_set_column_width(root: ET.Element, address: str, width: float) -> Dict[str, Any]:
    if width <= 0 or width > 255:
        raise OfficePackageError("Excel column width must be greater than 0 and at most 255.")
    endpoints = str(address).upper().split(":")
    start_column = excel_address_parts(endpoints[0])[0]
    end_column = excel_address_parts(endpoints[-1])[0]
    columns = root.find(q("s", "cols"))
    if columns is None:
        columns = ET.Element(q("s", "cols"))
        sheet_data = root.find(q("s", "sheetData"))
        root.insert(list(root).index(sheet_data) if sheet_data is not None else 0, columns)
    for column_number in range(start_column, end_column + 1):
        for existing in list(columns):
            minimum = int(existing.attrib.get("min", "0"))
            maximum = int(existing.attrib.get("max", "0"))
            if minimum <= column_number <= maximum:
                columns.remove(existing)
                if minimum < column_number:
                    columns.append(ET.Element(q("s", "col"), {**existing.attrib, "max": str(column_number - 1)}))
                if column_number < maximum:
                    columns.append(ET.Element(q("s", "col"), {**existing.attrib, "min": str(column_number + 1)}))
        columns.append(ET.Element(q("s", "col"), {"min": str(column_number), "max": str(column_number), "width": str(width), "customWidth": "1"}))
    return {"startColumn": start_column, "endColumn": end_column, "width": width}


def office_relationship_part(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def excel_relationship_root(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], part: str) -> Tuple[str, ET.Element]:
    relationship_part = office_relationship_part(part)
    if modifications.get(relationship_part) is not None:
        return relationship_part, ET.fromstring(modifications[relationship_part])
    if relationship_part in package.namelist():
        return relationship_part, ET.fromstring(package.read(relationship_part))
    return relationship_part, ET.Element(q("pr", "Relationships"))


def excel_add_relationship(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], part: str, rel_type: str, target: str) -> str:
    relationship_part, root = excel_relationship_root(package, modifications, part)
    used = {entry.attrib.get("Id") for entry in root.findall(q("pr", "Relationship"))}
    number = 1
    while "rId%d" % number in used:
        number += 1
    rel_id = "rId%d" % number
    ET.SubElement(root, q("pr", "Relationship"), {"Id": rel_id, "Type": rel_type, "Target": target})
    modifications[relationship_part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return rel_id


def excel_next_part_number(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], pattern: str) -> int:
    numbers = [int(match.group(1)) for name in set(package.namelist()) | set(modifications) if (match := re.fullmatch(pattern, name)) and modifications.get(name, b"present") is not None]
    return max(numbers or [0]) + 1


def excel_source_range(sheet_name: str, source_data: str) -> Tuple[str, List[List[str]]]:
    source = str(source_data or "")
    if "!" in source:
        supplied_sheet, address = source.rsplit("!", 1)
        supplied_sheet = supplied_sheet.strip("'").replace("''", "'")
        if supplied_sheet != sheet_name:
            raise OfficePackageError("OpenXML chart sourceData must reference the selected worksheet.")
    else:
        address = source
    addresses = excel_range_addresses(address.replace("$", ""))
    if len(addresses) < 2 or len(addresses[0]) < 2:
        raise OfficePackageError("Excel chart sourceData must include a header row, category column, and at least one data row.")
    escaped_sheet = "'%s'" % sheet_name.replace("'", "''")
    return escaped_sheet, addresses


def excel_chart_xml(sheet_name: str, source_data: str, chart_type: str, title: Optional[str]) -> bytes:
    sheet_formula, addresses = excel_source_range(sheet_name, source_data)
    normalized = re.sub(r"[^a-z]", "", str(chart_type or "").lower())
    if normalized in {"bar", "barclustered"}:
        element_name, bar_direction = "barChart", "bar"
    elif normalized in {"column", "columnclustered", "clusteredcolumn"}:
        element_name, bar_direction = "barChart", "col"
    elif normalized in {"line", "linechart"}:
        element_name, bar_direction = "lineChart", None
    elif normalized in {"pie", "piechart"}:
        element_name, bar_direction = "pieChart", None
    else:
        raise OfficePackageError("OpenXML chartType must be BarClustered, ColumnClustered, Line, or Pie.")
    root = ET.Element(q("c", "chartSpace"))
    chart = ET.SubElement(root, q("c", "chart"))
    if title is not None:
        title_node = ET.SubElement(chart, q("c", "title"))
        tx = ET.SubElement(title_node, q("c", "tx"))
        rich = ET.SubElement(tx, q("c", "rich"))
        ET.SubElement(rich, q("a", "bodyPr")); ET.SubElement(rich, q("a", "lstStyle"))
        paragraph = ET.SubElement(rich, q("a", "p")); run = ET.SubElement(paragraph, q("a", "r")); ET.SubElement(run, q("a", "t")).text = str(title)
    plot = ET.SubElement(chart, q("c", "plotArea")); ET.SubElement(plot, q("c", "layout"))
    chart_element = ET.SubElement(plot, q("c", element_name))
    if bar_direction:
        ET.SubElement(chart_element, q("c", "barDir"), {"val": bar_direction})
        ET.SubElement(chart_element, q("c", "grouping"), {"val": "clustered"})
    start = excel_address_parts(addresses[0][0]); end = excel_address_parts(addresses[-1][-1])
    category_start = "%s%d" % (excel_column_name(start[0]), start[1] + 1)
    category_end = "%s%d" % (excel_column_name(start[0]), end[1])
    category_formula = "%s!$%s$%d:$%s$%d" % (sheet_formula, excel_column_name(start[0]), start[1] + 1, excel_column_name(start[0]), end[1])
    for series_index, column in enumerate(range(start[0] + 1, end[0] + 1)):
        series = ET.SubElement(chart_element, q("c", "ser"))
        ET.SubElement(series, q("c", "idx"), {"val": str(series_index)}); ET.SubElement(series, q("c", "order"), {"val": str(series_index)})
        tx = ET.SubElement(series, q("c", "tx")); ref = ET.SubElement(tx, q("c", "strRef")); ET.SubElement(ref, q("c", "f")).text = "%s!$%s$%d" % (sheet_formula, excel_column_name(column), start[1])
        cat = ET.SubElement(series, q("c", "cat")); cat_ref = ET.SubElement(cat, q("c", "strRef")); ET.SubElement(cat_ref, q("c", "f")).text = category_formula
        values = ET.SubElement(series, q("c", "val")); number_ref = ET.SubElement(values, q("c", "numRef")); ET.SubElement(number_ref, q("c", "f")).text = "%s!$%s$%d:$%s$%d" % (sheet_formula, excel_column_name(column), start[1] + 1, excel_column_name(column), end[1])
    if element_name != "pieChart":
        category_axis_id, value_axis_id = "123456", "123457"
        ET.SubElement(chart_element, q("c", "axId"), {"val": category_axis_id}); ET.SubElement(chart_element, q("c", "axId"), {"val": value_axis_id})
        category_axis = ET.SubElement(plot, q("c", "catAx")); ET.SubElement(category_axis, q("c", "axId"), {"val": category_axis_id}); ET.SubElement(category_axis, q("c", "axPos"), {"val": "b"}); ET.SubElement(category_axis, q("c", "crossAx"), {"val": value_axis_id}); ET.SubElement(category_axis, q("c", "crosses"), {"val": "autoZero"})
        value_axis = ET.SubElement(plot, q("c", "valAx")); ET.SubElement(value_axis, q("c", "axId"), {"val": value_axis_id}); ET.SubElement(value_axis, q("c", "axPos"), {"val": "l"}); ET.SubElement(value_axis, q("c", "crossAx"), {"val": category_axis_id}); ET.SubElement(value_axis, q("c", "crosses"), {"val": "autoZero"})
    ET.SubElement(chart, q("c", "plotVisOnly"), {"val": "1"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def excel_set_chart_title(root: ET.Element, title: str) -> None:
    chart = root.find(q("c", "chart"))
    if chart is None:
        raise OfficePackageError("Excel chart part has no chart element.")
    existing = chart.find(q("c", "title"))
    if existing is not None:
        chart.remove(existing)
    title_node = ET.Element(q("c", "title")); tx = ET.SubElement(title_node, q("c", "tx")); rich = ET.SubElement(tx, q("c", "rich"))
    ET.SubElement(rich, q("a", "bodyPr")); ET.SubElement(rich, q("a", "lstStyle")); paragraph = ET.SubElement(rich, q("a", "p")); run = ET.SubElement(paragraph, q("a", "r")); ET.SubElement(run, q("a", "t")).text = str(title)
    chart.insert(0, title_node)


def excel_sheet_drawing(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], sheet_part: str, sheet_root: ET.Element) -> Tuple[str, ET.Element]:
    relationship_part, relationships = excel_relationship_root(package, modifications, sheet_part)
    drawing_relationship = next((entry for entry in relationships.findall(q("pr", "Relationship")) if entry.attrib.get("Type", "").endswith("/drawing")), None)
    if drawing_relationship is not None:
        drawing_part = resolved_relationship_target(sheet_part, drawing_relationship.attrib.get("Target", ""))
        payload = modifications.get(drawing_part) if drawing_part in modifications else package.read(drawing_part)
        return drawing_part, ET.fromstring(payload)
    number = excel_next_part_number(package, modifications, r"xl/drawings/drawing(\d+)\.xml")
    drawing_part = "xl/drawings/drawing%d.xml" % number
    target = posixpath.relpath(drawing_part, posixpath.dirname(sheet_part))
    rel_id = excel_add_relationship(package, modifications, sheet_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", target)
    ET.SubElement(sheet_root, q("s", "drawing"), {q("r", "id"): rel_id})
    add_content_type_override(package, modifications, "/" + drawing_part, "application/vnd.openxmlformats-officedocument.drawing+xml")
    return drawing_part, ET.Element(q("xdr", "wsDr"))


def excel_create_chart(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], sheet_part: str, sheet_root: ET.Element, operation: Dict[str, Any]) -> Dict[str, Any]:
    drawing_part, drawing = excel_sheet_drawing(package, modifications, sheet_part, sheet_root)
    chart_number = excel_next_part_number(package, modifications, r"xl/charts/chart(\d+)\.xml")
    chart_part = "xl/charts/chart%d.xml" % chart_number
    chart_target = posixpath.relpath(chart_part, posixpath.dirname(drawing_part))
    rel_id = excel_add_relationship(package, modifications, drawing_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", chart_target)
    used_ids = [int(node.attrib.get("id", "0")) for node in drawing.iter(q("xdr", "cNvPr")) if node.attrib.get("id", "").isdigit()]
    frame_id = max(used_ids or [0]) + 1
    name = str(operation.get("name") or "Chart %d" % chart_number)
    anchor = ET.SubElement(drawing, q("xdr", "oneCellAnchor"))
    start = ET.SubElement(anchor, q("xdr", "from"))
    ET.SubElement(start, q("xdr", "col")).text = "0"; ET.SubElement(start, q("xdr", "colOff")).text = str(round(float(operation.get("left", 0)) * 12700))
    ET.SubElement(start, q("xdr", "row")).text = "0"; ET.SubElement(start, q("xdr", "rowOff")).text = str(round(float(operation.get("top", 0)) * 12700))
    ET.SubElement(anchor, q("xdr", "ext"), {"cx": str(round(float(operation.get("width", 480)) * 12700)), "cy": str(round(float(operation.get("height", 288)) * 12700))})
    frame = ET.SubElement(anchor, q("xdr", "graphicFrame"), {"macro": ""})
    nv = ET.SubElement(frame, q("xdr", "nvGraphicFramePr")); ET.SubElement(nv, q("xdr", "cNvPr"), {"id": str(frame_id), "name": name}); ET.SubElement(nv, q("xdr", "cNvGraphicFramePr"))
    transform = ET.SubElement(frame, q("xdr", "xfrm")); ET.SubElement(transform, q("a", "off"), {"x": "0", "y": "0"}); ET.SubElement(transform, q("a", "ext"), {"cx": "0", "cy": "0"})
    graphic = ET.SubElement(frame, q("a", "graphic")); data = ET.SubElement(graphic, q("a", "graphicData"), {"uri": NS["c"]}); ET.SubElement(data, q("c", "chart"), {q("r", "id"): rel_id})
    ET.SubElement(anchor, q("xdr", "clientData"))
    modifications[drawing_part] = ET.tostring(drawing, encoding="utf-8", xml_declaration=True)
    modifications[chart_part] = excel_chart_xml(str(operation.get("sheet", "")), str(operation.get("sourceData", "")), str(operation.get("chartType", "")), operation.get("titleText"))
    add_content_type_override(package, modifications, "/" + chart_part, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
    return {"name": name, "part": chart_part, "type": operation.get("chartType"), "sourceData": operation.get("sourceData"), "relationshipId": rel_id}


def excel_update_chart(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], sheet_part: str, operation: Dict[str, Any]) -> Dict[str, Any]:
    requested = str(operation.get("chart", ""))
    _, sheet_relationships = excel_relationship_root(package, modifications, sheet_part)
    for relationship_node in sheet_relationships.findall(q("pr", "Relationship")):
        relationship = relationship_node.attrib
        if not relationship.get("Type", "").endswith("/drawing"):
            continue
        drawing_part = resolved_relationship_target(sheet_part, relationship.get("Target", ""))
        drawing = ET.fromstring(modifications.get(drawing_part) or package.read(drawing_part))
        drawing_relationship_part, drawing_relationships = excel_relationship_root(package, modifications, drawing_part)
        relationship_map = {entry.attrib.get("Id"): entry for entry in drawing_relationships.findall(q("pr", "Relationship"))}
        for frame in drawing.iter(q("xdr", "graphicFrame")):
            name_node = frame.find(".//" + q("xdr", "cNvPr")); chart_node = frame.find(".//" + q("c", "chart"))
            if chart_node is None or name_node is None or requested not in {name_node.attrib.get("name"), name_node.attrib.get("id")}:
                continue
            chart_relationship = relationship_map.get(chart_node.attrib.get(q("r", "id")))
            if chart_relationship is None:
                raise OfficePackageError("Excel chart relationship is missing.")
            chart_part = resolved_relationship_target(drawing_part, chart_relationship.attrib.get("Target", ""))
            existing_root = ET.fromstring(modifications.get(chart_part) or package.read(chart_part))
            existing_type = next((node.tag.split("}")[-1] for node in existing_root.findall(".//" + q("c", "plotArea") + "/*") if node.tag.endswith("Chart")), "barChart")
            inferred_type = operation.get("chartType") or {"barChart": "ColumnClustered", "lineChart": "Line", "pieChart": "Pie"}.get(existing_type, "ColumnClustered")
            source_data = operation.get("sourceData")
            title = operation.get("titleText")
            if title is None:
                title_node = existing_root.find(".//" + q("c", "title")); title = text_of(title_node, (q("a", "t"), q("c", "v"))) if title_node is not None else None
            if source_data is not None:
                modifications[chart_part] = excel_chart_xml(str(operation.get("sheet", "")), str(source_data), str(inferred_type), title)
            elif operation.get("chartType") is not None:
                raise OfficePackageError("OpenXML chartType updates require sourceData so series can be rebuilt safely.")
            elif operation.get("titleText") is not None:
                excel_set_chart_title(existing_root, str(operation.get("titleText")))
                modifications[chart_part] = ET.tostring(existing_root, encoding="utf-8", xml_declaration=True)
            if operation.get("name") is not None:
                name_node.attrib["name"] = str(operation["name"])
            anchor = direct_parent(drawing, frame)
            start = anchor.find(q("xdr", "from")) if anchor is not None else None; extent = anchor.find(q("xdr", "ext")) if anchor is not None else None
            if start is not None:
                if operation.get("left") is not None: start.find(q("xdr", "colOff")).text = str(round(float(operation["left"]) * 12700))
                if operation.get("top") is not None: start.find(q("xdr", "rowOff")).text = str(round(float(operation["top"]) * 12700))
            if extent is not None:
                if operation.get("width") is not None: extent.attrib["cx"] = str(round(float(operation["width"]) * 12700))
                if operation.get("height") is not None: extent.attrib["cy"] = str(round(float(operation["height"]) * 12700))
            modifications[drawing_part] = ET.tostring(drawing, encoding="utf-8", xml_declaration=True)
            return {"name": name_node.attrib.get("name"), "part": chart_part, "type": inferred_type, "sourceData": source_data, "title": title}
    raise OfficePackageError("Excel chart was not found: %s" % requested)


def edit_excel(package: zipfile.ZipFile, request: Dict[str, Any]) -> Tuple[Dict[str, Optional[bytes]], List[Dict[str, Any]]]:
    sheet_parts = excel_sheet_parts(package)
    modifications: Dict[str, Optional[bytes]] = {}
    changes = []
    for operation in request.get("operations", []):
        op_type = operation.get("type")
        if op_type == "addTableRow":
            changes.append({"operation": op_type, **excel_add_table_rows(package, modifications, sheet_parts, operation)})
            continue
        if op_type == "setTableTotals":
            changes.append({"operation": op_type, **excel_set_table_totals(package, modifications, sheet_parts, operation)})
            continue
        if op_type == "recalculate":
            changes.append({"operation": op_type, **apply_excel_recalculate(package, modifications)})
            continue
        if op_type == "renameSheet":
            old_name = str(operation.get("sheet", ""))
            new_name = str(operation.get("newName", ""))
            if not new_name or len(new_name) > 31 or re.search(r"[\\/*?:\[\]]", new_name):
                raise OfficePackageError("Excel newName is not a valid worksheet name.")
            workbook = ET.fromstring(modifications.get("xl/workbook.xml", package.read("xl/workbook.xml")))
            sheets = list(workbook.find(q("s", "sheets")) or [])
            if any(sheet.attrib.get("name", "").lower() == new_name.lower() for sheet in sheets):
                raise OfficePackageError("Excel worksheet name already exists: %s" % new_name)
            target = next((sheet for sheet in sheets if sheet.attrib.get("name") == old_name), None)
            if target is None:
                raise OfficePackageError("Excel worksheet was not found: %s" % old_name)
            target.attrib["name"] = new_name
            modifications["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
            sheet_parts[new_name] = sheet_parts.pop(old_name)
            changes.append({"operation": op_type, "before": old_name, "after": new_name})
            continue
        if op_type == "setDefinedName":
            name = str(operation.get("name", ""))
            if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", name):
                raise OfficePackageError("Excel defined name is invalid.")
            workbook = ET.fromstring(modifications.get("xl/workbook.xml", package.read("xl/workbook.xml")))
            container = workbook.find(q("s", "definedNames"))
            if container is None:
                container = ET.SubElement(workbook, q("s", "definedNames"))
            target = next((entry for entry in list(container) if entry.attrib.get("name") == name), None)
            before = target.text if target is not None else None
            if target is None:
                target = ET.SubElement(container, q("s", "definedName"), {"name": name})
            target.text = str(operation.get("formula", ""))
            modifications["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
            changes.append({"operation": op_type, "name": name, "before": before, "after": target.text})
            continue
        if op_type not in {"setCell", "setFormula", "setRange", "clearRange", "setStyle", "setNumberFormat", "addConditionalFormat", "setDataValidation", "freezePanes", "setColumnWidth", "createChart", "updateChart"}:
            raise OfficePackageError("Unsupported Excel operation: %s" % op_type)
        sheet_name = str(operation.get("sheet", ""))
        part = sheet_parts.get(sheet_name)
        if not part:
            raise OfficePackageError("Excel worksheet was not found: %s" % sheet_name)
        address = str(operation.get("address", "")).upper()
        root = ET.fromstring(modifications.get(part, package.read(part)))
        shared = excel_shared_strings(package)
        if op_type == "createChart":
            after_chart = excel_create_chart(package, modifications, part, root, operation)
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changes.append({"operation": op_type, "sheet": sheet_name, "after": after_chart})
            continue
        if op_type == "updateChart":
            after_chart = excel_update_chart(package, modifications, part, operation)
            changes.append({"operation": op_type, "sheet": sheet_name, "chart": operation.get("chart"), "after": after_chart})
            continue
        addresses = [] if op_type == "freezePanes" else excel_range_addresses(address)
        if op_type == "addConditionalFormat":
            rule_type = str(operation.get("ruleType", "expression"))
            if rule_type not in {"expression", "cellIs"}:
                raise OfficePackageError("Excel conditional-format ruleType must be expression or cellIs.")
            formula = str(operation.get("formula", ""))
            if not formula:
                raise OfficePackageError("Excel conditional formatting requires formula.")
            dxf_id = excel_add_fill_dxf(package, modifications, str(operation.get("fillColor", "")))
            priority = 1 + max([int(rule.attrib.get("priority", "0")) for rule in root.iter(q("s", "cfRule"))] or [0])
            container = ET.SubElement(root, q("s", "conditionalFormatting"), {"sqref": address})
            attributes = {"type": rule_type, "dxfId": str(dxf_id), "priority": str(priority)}
            if rule_type == "cellIs":
                operator = str(operation.get("operator", "equal"))
                if operator not in {"between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"}:
                    raise OfficePackageError("Excel conditional-format operator is invalid.")
                attributes["operator"] = operator
            rule = ET.SubElement(container, q("s", "cfRule"), attributes)
            ET.SubElement(rule, q("s", "formula")).text = formula.lstrip("=")
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changes.append({"operation": op_type, "sheet": sheet_name, "address": address, "after": {"ruleType": rule_type, "formula": formula, "fillColor": operation.get("fillColor"), "priority": priority}})
            continue
        if op_type == "setDataValidation":
            validation_type = str(operation.get("validationType", "custom"))
            if validation_type not in {"whole", "decimal", "list", "date", "time", "textLength", "custom"}:
                raise OfficePackageError("Excel data-validation type is invalid.")
            validations = root.find(q("s", "dataValidations"))
            if validations is None:
                validations = ET.SubElement(root, q("s", "dataValidations"), {"count": "0"})
            for existing in list(validations):
                if existing.attrib.get("sqref") == address:
                    validations.remove(existing)
            attributes = {"type": validation_type, "sqref": address, "allowBlank": "1" if operation.get("allowBlank", True) else "0"}
            if operation.get("operator"):
                attributes["operator"] = str(operation["operator"])
            validation = ET.SubElement(validations, q("s", "dataValidation"), attributes)
            for field in ("formula1", "formula2"):
                if operation.get(field) is not None:
                    ET.SubElement(validation, q("s", field)).text = str(operation[field])
            validations.attrib["count"] = str(len(list(validations)))
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changes.append({"operation": op_type, "sheet": sheet_name, "address": address, "after": {"validationType": validation_type, "formula1": operation.get("formula1"), "formula2": operation.get("formula2")}})
            continue
        if op_type == "freezePanes":
            rows = int(operation.get("rows", 0))
            columns_count = int(operation.get("columns", 0))
            if rows < 0 or columns_count < 0 or rows > 1_048_575 or columns_count > 16_383:
                raise OfficePackageError("Excel freeze pane counts are outside worksheet limits.")
            views = root.find(q("s", "sheetViews"))
            if views is None:
                views = ET.Element(q("s", "sheetViews"))
                root.insert(0, views)
            view = views.find(q("s", "sheetView"))
            if view is None:
                view = ET.SubElement(views, q("s", "sheetView"), {"workbookViewId": "0"})
            existing = view.find(q("s", "pane"))
            if existing is not None:
                view.remove(existing)
            if rows or columns_count:
                top_left = "%s%d" % (excel_column_name(columns_count + 1), rows + 1)
                pane = {"state": "frozen", "topLeftCell": top_left, "activePane": "bottomRight" if rows and columns_count else "bottomLeft" if rows else "topRight"}
                if rows: pane["ySplit"] = str(rows)
                if columns_count: pane["xSplit"] = str(columns_count)
                view.insert(0, ET.Element(q("s", "pane"), pane))
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changes.append({"operation": op_type, "sheet": sheet_name, "after": {"rows": rows, "columns": columns_count}})
            continue
        if op_type == "setColumnWidth":
            after_width = excel_set_column_width(root, address, float(operation.get("width", 0)))
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changes.append({"operation": op_type, "sheet": sheet_name, "address": address, "after": after_width})
            continue
        values = operation.get("values")
        formulas = operation.get("formulas")
        if op_type == "setRange":
            if values is None and formulas is None:
                raise OfficePackageError("Excel setRange requires values or formulas.")
            for matrix, label in ((values, "values"), (formulas, "formulas")):
                if matrix is not None and (not isinstance(matrix, list) or len(matrix) != len(addresses) or any(not isinstance(row, list) or len(row) != len(addresses[0]) for row in matrix)):
                    raise OfficePackageError("Excel setRange.%s dimensions must match address." % label)
        for row_index, row_addresses in enumerate(addresses):
            for column_index, cell_address in enumerate(row_addresses):
                cell = find_or_create_excel_cell(root, cell_address)
                formula_node = cell.find(q("s", "f"))
                before = {"value": excel_cell_value(cell, shared), "formula": formula_node.text if formula_node is not None else None, "styleIndex": int(cell.attrib.get("s", "0"))}
                if op_type == "setFormula":
                    set_excel_cell(cell, operation.get("value"), operation.get("formula"))
                elif op_type == "setCell":
                    set_excel_cell(cell, operation.get("value"))
                elif op_type == "setRange":
                    value = values[row_index][column_index] if values is not None else None
                    formula = formulas[row_index][column_index] if formulas is not None else None
                    set_excel_cell(cell, value, formula)
                elif op_type == "clearRange":
                    if operation.get("contents", True):
                        set_excel_cell(cell, None)
                    if operation.get("format", False):
                        cell.attrib.pop("s", None)
                else:
                    if op_type == "setStyle":
                        style_index = int(operation.get("styleIndex", -1))
                        if style_index < 0:
                            raise OfficePackageError("Excel styleIndex must be non-negative.")
                        cell.attrib["s"] = str(style_index)
                    else:
                        cell.attrib["s"] = str(excel_style_with_number_format(
                            package,
                            modifications,
                            int(cell.attrib.get("s", "0")),
                            str(operation.get("formatCode", "")),
                        ))
                after_formula = cell.find(q("s", "f"))
                after = {"value": excel_cell_value(cell, []), "formula": after_formula.text if after_formula is not None else None, "styleIndex": int(cell.attrib.get("s", "0"))}
                if op_type == "setNumberFormat":
                    after["numberFormat"] = operation.get("formatCode")
                changes.append({"operation": op_type, "sheet": sheet_name, "address": cell_address, "before": before, "after": after})
        modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return modifications, changes


def ppt_current_slides(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]]) -> List[Dict[str, Any]]:
    presentation = ET.fromstring(modifications.get("ppt/presentation.xml") or package.read("ppt/presentation.xml"))
    rels_name = "ppt/_rels/presentation.xml.rels"
    rels_root = ET.fromstring(modifications.get(rels_name) or package.read(rels_name))
    relationships = {rel.attrib.get("Id"): rel for rel in rels_root.findall(q("pr", "Relationship"))}
    result = []
    for index, node in enumerate(list(presentation.find(q("p", "sldIdLst")) or [])):
        rel_id = node.attrib.get(q("r", "id"))
        rel = relationships.get(rel_id)
        result.append({"index": index, "node": node, "relationship": rel, "relationshipId": rel_id, "part": resolved_relationship_target("ppt/presentation.xml", rel.attrib.get("Target", "")) if rel is not None else None})
    return result


PPT_IMAGE_CONTENT_TYPES = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
}


def ppt_shape_by_id(root: ET.Element, shape_id: str) -> Optional[ET.Element]:
    for tag in (q("p", "sp"), q("p", "pic"), q("p", "graphicFrame"), q("p", "grpSp")):
        for shape in root.iter(tag):
            nv = shape.find(".//" + q("p", "cNvPr"))
            if nv is not None and nv.attrib.get("id") == shape_id:
                return shape
    return None


def ppt_add_text_box(root: ET.Element, operation: Dict[str, Any]) -> Dict[str, Any]:
    shape_tree = root.find(".//" + q("p", "spTree"))
    if shape_tree is None:
        raise OfficePackageError("PowerPoint slide shape tree is missing.")
    used_ids = {int(node.attrib["id"]) for node in root.iter(q("p", "cNvPr")) if node.attrib.get("id", "").isdigit()}
    requested_id = operation.get("shapeId")
    shape_id = int(requested_id) if requested_id is not None else max(used_ids or {1}) + 1
    if shape_id < 1 or shape_id in used_ids:
        raise OfficePackageError("PowerPoint addTextBox shapeId must be a unique positive integer on the slide.")
    x = int(operation.get("x", 0))
    y = int(operation.get("y", 0))
    width = int(operation.get("width", 1_828_800))
    height = int(operation.get("height", 457_200))
    if width <= 0 or height <= 0:
        raise OfficePackageError("PowerPoint addTextBox width and height must be positive.")
    name = str(operation.get("name") or "TextBox %d" % shape_id)
    shape = ET.Element(q("p", "sp"))
    non_visual = ET.SubElement(shape, q("p", "nvSpPr"))
    ET.SubElement(non_visual, q("p", "cNvPr"), {"id": str(shape_id), "name": name})
    ET.SubElement(non_visual, q("p", "cNvSpPr"), {"txBox": "1"})
    ET.SubElement(non_visual, q("p", "nvPr"))
    shape_properties = ET.SubElement(shape, q("p", "spPr"))
    transform = ET.SubElement(shape_properties, q("a", "xfrm"))
    ET.SubElement(transform, q("a", "off"), {"x": str(x), "y": str(y)})
    ET.SubElement(transform, q("a", "ext"), {"cx": str(width), "cy": str(height)})
    geometry = ET.SubElement(shape_properties, q("a", "prstGeom"), {"prst": "rect"})
    ET.SubElement(geometry, q("a", "avLst"))
    ET.SubElement(shape_properties, q("a", "noFill"))
    line = ET.SubElement(shape_properties, q("a", "ln"))
    ET.SubElement(line, q("a", "noFill"))
    text_body = ET.SubElement(shape, q("p", "txBody"))
    ET.SubElement(text_body, q("a", "bodyPr"), {"wrap": "square"})
    ET.SubElement(text_body, q("a", "lstStyle"))
    paragraph = ET.SubElement(text_body, q("a", "p"))
    run = ET.SubElement(paragraph, q("a", "r"))
    ET.SubElement(run, q("a", "rPr"), {"lang": "en-US"})
    text_node = ET.SubElement(run, q("a", "t"))
    text_node.text = str(operation.get("text", ""))
    if text_node.text[:1].isspace() or text_node.text[-1:].isspace():
        text_node.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
    ET.SubElement(paragraph, q("a", "endParaRPr"), {"lang": "en-US"})
    shape_tree.append(shape)
    return {"shapeId": str(shape_id), "name": name, "text": str(operation.get("text", "")), "geometry": {"x": x, "y": y, "width": width, "height": height}}


def ppt_set_text_style(shape: ET.Element, operation: Dict[str, Any]) -> Dict[str, Any]:
    style_keys = {"fontFamily", "fontSize", "bold", "italic", "underline", "color"}
    if not any(key in operation for key in style_keys):
        raise OfficePackageError("PowerPoint setTextStyle requires at least one style property.")
    runs = [node for node in shape.iter() if node.tag in {q("a", "r"), q("a", "fld")}]
    if not runs:
        raise OfficePackageError("Selected PowerPoint shape has no editable text runs.")
    color = operation.get("color")
    if color is not None and not re.fullmatch(r"[0-9A-Fa-f]{6}", str(color)):
        raise OfficePackageError("PowerPoint setTextStyle.color must be a six-digit RGB hex value.")
    font_size = operation.get("fontSize")
    if font_size is not None and not (1 <= float(font_size) <= 400):
        raise OfficePackageError("PowerPoint setTextStyle.fontSize must be between 1 and 400 points.")
    for run in runs:
        props = run.find(q("a", "rPr"))
        if props is None:
            props = ET.Element(q("a", "rPr"))
            run.insert(0, props)
        if font_size is not None:
            props.attrib["sz"] = str(round(float(font_size) * 100))
        if operation.get("bold") is not None:
            props.attrib["b"] = "1" if operation["bold"] else "0"
        if operation.get("italic") is not None:
            props.attrib["i"] = "1" if operation["italic"] else "0"
        if operation.get("underline") is not None:
            props.attrib["u"] = "sng" if operation["underline"] else "none"
        if operation.get("fontFamily") is not None:
            latin = props.find(q("a", "latin"))
            if latin is None:
                latin = ET.SubElement(props, q("a", "latin"))
            latin.attrib["typeface"] = str(operation["fontFamily"])
        if color is not None:
            for child in list(props):
                if child.tag in {q("a", "solidFill"), q("a", "gradFill"), q("a", "noFill"), q("a", "pattFill")}:
                    props.remove(child)
            fill = ET.Element(q("a", "solidFill"))
            ET.SubElement(fill, q("a", "srgbClr"), {"val": str(color).upper()})
            typeface_tags = {q("a", "latin"), q("a", "ea"), q("a", "cs"), q("a", "sym"), q("a", "hlinkClick"), q("a", "hlinkMouseOver")}
            insert_at = next((index for index, child in enumerate(list(props)) if child.tag in typeface_tags), len(list(props)))
            props.insert(insert_at, fill)
    return {key: operation[key] for key in style_keys if key in operation}


def ppt_replace_image(package: zipfile.ZipFile, modifications: Dict[str, Optional[bytes]], slide_part: str, slide_root: ET.Element, shape: ET.Element, operation: Dict[str, Any]) -> Dict[str, Any]:
    if shape.tag != q("p", "pic"):
        raise OfficePackageError("PowerPoint replaceImage requires a picture shape.")
    content_type = str(operation.get("contentType", "")).lower()
    extension = PPT_IMAGE_CONTENT_TYPES.get(content_type)
    if not extension:
        raise OfficePackageError("PowerPoint replaceImage contentType is unsupported.")
    try:
        payload = base64.b64decode(str(operation.get("base64", "")), validate=True)
    except (binascii.Error, ValueError) as error:
        raise OfficePackageError("PowerPoint replaceImage.base64 is invalid.") from error
    if not payload or len(payload) > 25 * 1024 * 1024:
        raise OfficePackageError("PowerPoint replacement image must be between 1 byte and 25 MiB.")
    signatures_match = {
        "image/png": payload.startswith(b"\x89PNG\r\n\x1a\n"),
        "image/jpeg": payload.startswith(b"\xff\xd8\xff"),
        "image/gif": payload.startswith((b"GIF87a", b"GIF89a")),
        "image/bmp": payload.startswith(b"BM"),
        "image/tiff": payload.startswith((b"II*\x00", b"MM\x00*")),
    }
    if not signatures_match[content_type]:
        raise OfficePackageError("PowerPoint replacement image bytes do not match contentType.")
    blip = shape.find(".//" + q("a", "blip"))
    if blip is None:
        raise OfficePackageError("Selected PowerPoint picture has no embedded image reference.")
    existing_numbers = [int(match.group(1)) for name in set(package.namelist()) | set(modifications) if (match := re.fullmatch(r"ppt/media/image(\d+)\.[A-Za-z0-9]+", name))]
    media_part = "ppt/media/image%d.%s" % (max(existing_numbers or [0]) + 1, extension)
    modifications[media_part] = payload
    rels_name = posixpath.join(posixpath.dirname(slide_part), "_rels", posixpath.basename(slide_part) + ".rels")
    if rels_name in modifications and modifications[rels_name] is not None:
        rels_root = ET.fromstring(modifications[rels_name])
    elif rels_name in package.namelist():
        rels_root = ET.fromstring(package.read(rels_name))
    else:
        rels_root = ET.Element(q("pr", "Relationships"))
    used_rids = {entry.attrib.get("Id") for entry in rels_root.findall(q("pr", "Relationship"))}
    rid_number = 1
    while "rId%d" % rid_number in used_rids:
        rid_number += 1
    new_rid = "rId%d" % rid_number
    ET.SubElement(rels_root, q("pr", "Relationship"), {
        "Id": new_rid,
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "Target": "../media/%s" % posixpath.basename(media_part),
    })
    blip.attrib[q("r", "embed")] = new_rid
    ET.register_namespace("", NS["pr"])
    modifications[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    content_types_name = "[Content_Types].xml"
    if content_types_name not in package.namelist():
        raise OfficePackageError("PowerPoint package is missing [Content_Types].xml.")
    content_root = ET.fromstring(modifications.get(content_types_name) or package.read(content_types_name))
    defaults = [node for node in list(content_root) if node.tag == q("ct", "Default") and node.attrib.get("Extension", "").lower() == extension]
    if not any(node.attrib.get("ContentType") == content_type for node in defaults):
        if defaults:
            ET.SubElement(content_root, q("ct", "Override"), {"PartName": "/" + media_part, "ContentType": content_type})
        else:
            ET.SubElement(content_root, q("ct", "Default"), {"Extension": extension, "ContentType": content_type})
    ET.register_namespace("", NS["ct"])
    modifications[content_types_name] = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)
    return {"relationshipId": new_rid, "mediaPart": media_part, "contentType": content_type, "bytes": len(payload)}


def edit_powerpoint(package: zipfile.ZipFile, request: Dict[str, Any]) -> Tuple[Dict[str, Optional[bytes]], List[Dict[str, Any]]]:
    presentation = inspect_powerpoint(package, {"maxSlides": 5000})
    modifications: Dict[str, Optional[bytes]] = {}
    changes = []
    for operation in request.get("operations", []):
        op_type = operation.get("type")
        if op_type in {"moveSlide", "deleteSlide", "duplicateSlide"}:
            slides = ppt_current_slides(package, modifications)
            source_index = int(operation.get("slideIndex", -1))
            if source_index < 0 or source_index >= len(slides):
                raise OfficePackageError("PowerPoint slideIndex is out of range.")
            presentation_root = ET.fromstring(modifications.get("ppt/presentation.xml") or package.read("ppt/presentation.xml"))
            list_node = presentation_root.find(q("p", "sldIdLst"))
            rels_name = "ppt/_rels/presentation.xml.rels"
            rels_root = ET.fromstring(modifications.get(rels_name) or package.read(rels_name))
            source = slides[source_index]
            if op_type == "moveSlide":
                target_index = int(operation.get("toIndex", -1))
                if target_index < 0 or target_index >= len(slides):
                    raise OfficePackageError("PowerPoint toIndex is out of range.")
                node = list(list_node)[source_index]
                list_node.remove(node)
                list_node.insert(target_index, node)
                changes.append({"operation": op_type, "slideIndex": source_index, "toIndex": target_index})
            elif op_type == "deleteSlide":
                node = list(list_node)[source_index]
                list_node.remove(node)
                rel = next((entry for entry in rels_root.findall(q("pr", "Relationship")) if entry.attrib.get("Id") == source["relationshipId"]), None)
                if rel is not None:
                    rels_root.remove(rel)
                if source["part"]:
                    modifications[source["part"]] = None
                    source_rels = posixpath.join(posixpath.dirname(source["part"]), "_rels", posixpath.basename(source["part"]) + ".rels")
                    if source_rels in package.namelist():
                        modifications[source_rels] = None
                changes.append({"operation": op_type, "slideIndex": source_index, "part": source["part"]})
            else:
                source_part = source["part"]
                if not source_part:
                    raise OfficePackageError("PowerPoint source slide relationship is missing.")
                existing_numbers = [int(match.group(1)) for name in set(package.namelist()) | set(modifications) if (match := re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name))]
                new_part = "ppt/slides/slide%d.xml" % (max(existing_numbers or [0]) + 1)
                modifications[new_part] = modifications.get(source_part) or package.read(source_part)
                source_rels = posixpath.join(posixpath.dirname(source_part), "_rels", posixpath.basename(source_part) + ".rels")
                new_rels = posixpath.join(posixpath.dirname(new_part), "_rels", posixpath.basename(new_part) + ".rels")
                if source_rels in package.namelist() and not (source_rels in modifications and modifications[source_rels] is None):
                    modifications[new_rels] = modifications[source_rels] if source_rels in modifications else package.read(source_rels)
                used_rids = {entry.attrib.get("Id") for entry in rels_root.findall(q("pr", "Relationship"))}
                rid_number = 1
                while "rId%d" % rid_number in used_rids:
                    rid_number += 1
                new_rid = "rId%d" % rid_number
                ET.SubElement(rels_root, q("pr", "Relationship"), {"Id": new_rid, "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", "Target": "slides/%s" % posixpath.basename(new_part)})
                max_id = max([int(entry.attrib.get("id", "255")) for entry in list(list_node)] or [255])
                new_node = ET.Element(q("p", "sldId"), {"id": str(max_id + 1), q("r", "id"): new_rid})
                target_index = int(operation.get("toIndex", source_index + 1))
                if target_index < 0 or target_index > len(list(list_node)):
                    raise OfficePackageError("PowerPoint duplicate toIndex is out of range.")
                list_node.insert(target_index, new_node)
                if "[Content_Types].xml" in package.namelist():
                    content_root = ET.fromstring(modifications.get("[Content_Types].xml") or package.read("[Content_Types].xml"))
                    source_override = next((entry for entry in list(content_root) if entry.attrib.get("PartName") == "/" + source_part), None)
                    if source_override is not None:
                        ET.SubElement(content_root, source_override.tag, {**source_override.attrib, "PartName": "/" + new_part})
                        ET.register_namespace("", NS["ct"])
                        modifications["[Content_Types].xml"] = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)
                changes.append({"operation": op_type, "slideIndex": source_index, "toIndex": target_index, "part": new_part})
            modifications["ppt/presentation.xml"] = ET.tostring(presentation_root, encoding="utf-8", xml_declaration=True)
            ET.register_namespace("", NS["pr"])
            modifications[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
            continue
        if op_type not in {"replaceText", "setShapeText", "setShapeGeometry", "setTableCell", "setNotes", "addTextBox", "deleteShape", "setTextStyle", "replaceImage"}:
            raise OfficePackageError("Unsupported PowerPoint operation: %s" % op_type)
        requested_slide = operation.get("slideIndex")
        requested_shape = str(operation.get("shapeId")) if operation.get("shapeId") is not None else None
        operation_finished = False
        for slide in presentation["slides"]:
            if requested_slide is not None and slide["index"] != int(requested_slide):
                continue
            part = slide["part"]
            root = ET.fromstring(modifications.get(part, package.read(part)))
            if op_type == "addTextBox":
                created = ppt_add_text_box(root, operation)
                changes.append({"operation": op_type, "slideIndex": slide["index"], **created})
                modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                operation_finished = True
                break
            if op_type == "setNotes":
                notes_part = None
                for rel in rels_for_part(package, part).values():
                    if rel.get("Type", "").endswith("/notesSlide"):
                        notes_part = resolved_relationship_target(part, rel.get("Target", ""))
                        break
                if not notes_part or notes_part not in package.namelist():
                    raise OfficePackageError("PowerPoint slide has no editable notes part.")
                notes_root = ET.fromstring(modifications.get(notes_part, package.read(notes_part)))
                nodes = list(notes_root.iter(q("a", "t")))
                before = set_text_nodes(nodes, operation.get("text"))
                changes.append({"operation": op_type, "slideIndex": slide["index"], "before": before, "after": operation.get("text", "")})
                modifications[notes_part] = ET.tostring(notes_root, encoding="utf-8", xml_declaration=True)
                continue
            candidates = (root.findall(".//" + q("p", "sp"))
                          + root.findall(".//" + q("p", "pic"))
                          + root.findall(".//" + q("p", "graphicFrame"))
                          + root.findall(".//" + q("p", "grpSp")))
            for shape in candidates:
                nv = shape.find(".//" + q("p", "cNvPr"))
                shape_id = nv.attrib.get("id") if nv is not None else None
                if requested_shape is not None and shape_id != requested_shape:
                    continue
                if op_type == "deleteShape":
                    parent = direct_parent(root, shape)
                    if parent is None:
                        raise OfficePackageError("Could not locate the selected PowerPoint shape parent.")
                    before = ppt_shape_result(shape, slide["index"])
                    parent.remove(shape)
                    part_changes = [{"before": before, "after": None}]
                elif op_type == "setTextStyle":
                    before = {"runCount": len([node for node in shape.iter() if node.tag in {q("a", "r"), q("a", "fld")}])}
                    part_changes = [{"before": before, "after": ppt_set_text_style(shape, operation)}]
                elif op_type == "replaceImage":
                    before_blip = shape.find(".//" + q("a", "blip"))
                    before = {"relationshipId": before_blip.attrib.get(q("r", "embed")) if before_blip is not None else None}
                    part_changes = [{"before": before, "after": ppt_replace_image(package, modifications, part, root, shape, operation)}]
                elif op_type == "setShapeText":
                    before = set_text_nodes(list(shape.iter(q("a", "t"))), operation.get("text"))
                    part_changes = [{"before": before, "after": operation.get("text", "")}]
                elif op_type == "setShapeGeometry":
                    transform = shape.find(q("p", "xfrm"))
                    if transform is None:
                        transform = shape.find(".//" + q("a", "xfrm"))
                    if transform is None:
                        raise OfficePackageError("Selected PowerPoint shape has no editable transform.")
                    offset = transform.find(q("a", "off"))
                    extent = transform.find(q("a", "ext"))
                    before = {"x": offset.attrib.get("x") if offset is not None else None, "y": offset.attrib.get("y") if offset is not None else None, "width": extent.attrib.get("cx") if extent is not None else None, "height": extent.attrib.get("cy") if extent is not None else None}
                    if offset is not None:
                        if operation.get("x") is not None: offset.attrib["x"] = str(int(operation["x"]))
                        if operation.get("y") is not None: offset.attrib["y"] = str(int(operation["y"]))
                    if extent is not None:
                        if operation.get("width") is not None: extent.attrib["cx"] = str(int(operation["width"]))
                        if operation.get("height") is not None: extent.attrib["cy"] = str(int(operation["height"]))
                    part_changes = [{"before": before, "after": {key: operation.get(key) for key in ("x", "y", "width", "height") if operation.get(key) is not None}}]
                elif op_type == "setTableCell":
                    table = shape.find(".//" + q("a", "tbl"))
                    if table is None:
                        raise OfficePackageError("Selected PowerPoint shape is not a table.")
                    rows = table.findall(q("a", "tr"))
                    row_index = int(operation.get("rowIndex", -1))
                    column_index = int(operation.get("columnIndex", -1))
                    if row_index < 0 or row_index >= len(rows):
                        raise OfficePackageError("PowerPoint table rowIndex is out of range.")
                    cells = rows[row_index].findall(q("a", "tc"))
                    if column_index < 0 or column_index >= len(cells):
                        raise OfficePackageError("PowerPoint table columnIndex is out of range.")
                    before = set_text_nodes(list(cells[column_index].iter(q("a", "t"))), operation.get("text"))
                    part_changes = [{"before": before, "after": operation.get("text", ""), "rowIndex": row_index, "columnIndex": column_index}]
                else:
                    part_changes = replace_text_nodes(
                        list(shape.iter(q("a", "t"))),
                        str(operation.get("find", "")),
                        str(operation.get("replace", "")),
                        bool(operation.get("matchCase", True)),
                        bool(operation.get("all", True)),
                    )
                for change in part_changes:
                    changes.append({"operation": op_type, "slideIndex": slide["index"], "shapeId": shape_id, **change})
                if part_changes and op_type in {"deleteShape", "setTextStyle", "replaceImage"}:
                    operation_finished = True
                    break
                if part_changes and op_type == "replaceText" and not operation.get("all", True):
                    operation_finished = True
                    break
            modifications[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            if operation_finished:
                break
    return modifications, changes


def write_modified_package(source: Path, destination: Path, modifications: Dict[str, Optional[bytes]]) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as input_package, zipfile.ZipFile(destination, "w") as output_package:
        for info in input_package.infolist():
            if info.filename in modifications and modifications[info.filename] is None:
                continue
            payload = modifications.get(info.filename, input_package.read(info.filename))
            output_package.writestr(info, payload)
        existing = set(input_package.namelist())
        for name, payload in modifications.items():
            if name not in existing and payload is not None:
                output_package.writestr(name, payload)


def edit_package(path: Path, request: Dict[str, Any]) -> Dict[str, Any]:
    output_path = request.get("outputPath")
    if not isinstance(output_path, str) or not output_path:
        raise OfficePackageError("outputPath is required for edit.")
    destination = Path(output_path).expanduser().resolve()
    with zipfile.ZipFile(path, "r") as package:
        inventory = package_inventory(package)
        kind = detect_kind(package, inventory["names"])
        requested_kind = request.get("kind")
        if requested_kind and requested_kind != kind:
            raise OfficePackageError("Expected a %s package but detected %s." % (requested_kind, kind))
        if inventory["hasDigitalSignatures"]:
            raise OfficePackageError("Digitally signed Office packages are refused because editing would invalidate the signature. Remove the signature in Office before editing.")
        if inventory["hasMacros"] and not request.get("allowMacros", False):
            raise OfficePackageError("Macro-enabled Office packages require allowMacros=true and explicit confirmation.")
        if kind == "word":
            modifications, changes = edit_word(package, request)
        elif kind == "excel":
            modifications, changes = edit_excel(package, request)
        else:
            modifications, changes = edit_powerpoint(package, request)
    write_modified_package(path, destination, modifications)
    validation = validate_package(destination, {"kind": kind, "strictRelationships": True})
    return {"kind": kind, "outputPath": str(destination), "changes": changes, "changeCount": len(changes), "validation": validation}


def main() -> None:
    try:
        request = read_request()
        action = request.get("action")
        input_path = request.get("inputPath")
        if not isinstance(input_path, str) or not input_path:
            raise OfficePackageError("inputPath is required.")
        path = Path(input_path).expanduser().resolve()
        if action == "inspect":
            value = inspect_package(path, request)
        elif action == "validate":
            value = validate_package(path, request)
        elif action == "edit":
            value = edit_package(path, request)
        else:
            raise OfficePackageError("action must be inspect, validate, or edit.")
        print(json.dumps({"ok": True, "value": value}, ensure_ascii=False, separators=(",", ":")))
    except Exception as error:  # Keep subprocess failures structured for MCP callers.
        print(json.dumps({"ok": False, "error": str(error)}, ensure_ascii=False, separators=(",", ":")))
        sys.exit(1)


if __name__ == "__main__":
    main()
