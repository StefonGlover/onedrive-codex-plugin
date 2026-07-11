#!/usr/bin/env python3
"""Dependency-free Office Open XML inspection and validation helper.

The MCP server communicates with this process using one JSON request on stdin
and receives one JSON response on stdout.  The helper intentionally edits only
explicit package parts and copies every untouched ZIP member byte-for-byte.
"""

from __future__ import annotations

import hashlib
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET


MAX_PACKAGE_ENTRIES = 10_000
MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_COMPRESSION_RATIO = 250
MAX_TEXT_CHARS = 2_000_000

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
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


def inspect_excel(package: zipfile.ZipFile, request: Dict[str, Any]) -> Dict[str, Any]:
    workbook = xml_root(package, "xl/workbook.xml")
    relationships = rels_for_part(package, "xl/workbook.xml")
    shared = excel_shared_strings(package)
    max_cells = max(1, min(int(request.get("maxCells", 5000)), 50_000))
    include_cells = request.get("includeCells", True)
    sheets = []
    total_cells = 0
    sheet_nodes = workbook.find(q("s", "sheets"))
    for sheet in list(sheet_nodes or []):
        rel_id = sheet.attrib.get(q("r", "id"))
        relationship = relationships.get(rel_id or "", {})
        target = relationship.get("Target", "")
        part = resolved_relationship_target("xl/workbook.xml", target) if target else None
        sheet_result: Dict[str, Any] = {
            "name": sheet.attrib.get("name"),
            "sheetId": sheet.attrib.get("sheetId"),
            "relationshipId": rel_id,
            "state": sheet.attrib.get("state", "visible"),
            "part": part,
        }
        cells = []
        if include_cells and part and part in package.namelist():
            root = xml_root(package, part)
            for cell in root.iter(q("s", "c")):
                if total_cells >= max_cells:
                    break
                formula = cell.find(q("s", "f"))
                cells.append({
                    "address": cell.attrib.get("r"),
                    "value": excel_cell_value(cell, shared),
                    "formula": formula.text if formula is not None else None,
                    "styleIndex": int(cell.attrib.get("s", "0")),
                    "type": cell.attrib.get("t"),
                })
                total_cells += 1
        sheet_result["cells"] = cells
        sheet_result["cellCount"] = len(cells)
        sheets.append(sheet_result)
    defined_names = []
    names_node = workbook.find(q("s", "definedNames"))
    for name in list(names_node or []):
        defined_names.append({"name": name.attrib.get("name"), "value": name.text})
    return {
        "kind": "excel",
        "sheets": sheets,
        "sheetCount": len(sheets),
        "cellCount": total_cells,
        "definedNames": defined_names,
        "truncated": total_cells >= max_cells,
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
            if kind == "word":
                content = inspect_word(package, request)
            elif kind == "excel":
                content = inspect_excel(package, request)
            else:
                content = inspect_powerpoint(package, request)
            content.update({
                "package": {
                    "path": str(path),
                    "bytes": path.stat().st_size,
                    "entryCount": inventory["entryCount"],
                    "uncompressedBytes": inventory["uncompressedBytes"],
                    "hasMacros": inventory["hasMacros"],
                    "hasDigitalSignatures": inventory["hasDigitalSignatures"],
                    "relationshipErrors": relationship_errors,
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


def edit_word(package: zipfile.ZipFile, request: Dict[str, Any]) -> Tuple[Dict[str, bytes], List[Dict[str, Any]]]:
    names = set(package.namelist())
    modifications: Dict[str, bytes] = {}
    changes: List[Dict[str, Any]] = []
    for operation in request.get("operations", []):
        op_type = operation.get("type")
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


def edit_excel(package: zipfile.ZipFile, request: Dict[str, Any]) -> Tuple[Dict[str, bytes], List[Dict[str, Any]]]:
    sheet_parts = excel_sheet_parts(package)
    modifications: Dict[str, bytes] = {}
    changes = []
    for operation in request.get("operations", []):
        op_type = operation.get("type")
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
        if op_type not in {"setCell", "setFormula", "setRange", "clearRange", "setStyle"}:
            raise OfficePackageError("Unsupported Excel operation: %s" % op_type)
        sheet_name = str(operation.get("sheet", ""))
        part = sheet_parts.get(sheet_name)
        if not part:
            raise OfficePackageError("Excel worksheet was not found: %s" % sheet_name)
        address = str(operation.get("address", "")).upper()
        root = ET.fromstring(modifications.get(part, package.read(part)))
        shared = excel_shared_strings(package)
        addresses = excel_range_addresses(address)
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
                    style_index = int(operation.get("styleIndex", -1))
                    if style_index < 0:
                        raise OfficePackageError("Excel styleIndex must be non-negative.")
                    cell.attrib["s"] = str(style_index)
                after_formula = cell.find(q("s", "f"))
                changes.append({"operation": op_type, "sheet": sheet_name, "address": cell_address, "before": before, "after": {"value": excel_cell_value(cell, []), "formula": after_formula.text if after_formula is not None else None, "styleIndex": int(cell.attrib.get("s", "0"))}})
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
        if op_type not in {"replaceText", "setShapeText", "setShapeGeometry", "setTableCell", "setNotes"}:
            raise OfficePackageError("Unsupported PowerPoint operation: %s" % op_type)
        requested_slide = operation.get("slideIndex")
        requested_shape = str(operation.get("shapeId")) if operation.get("shapeId") is not None else None
        operation_finished = False
        for slide in presentation["slides"]:
            if requested_slide is not None and slide["index"] != int(requested_slide):
                continue
            part = slide["part"]
            root = ET.fromstring(modifications.get(part, package.read(part)))
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
            for shape in root.findall(".//" + q("p", "sp")) + root.findall(".//" + q("p", "graphicFrame")):
                nv = shape.find(".//" + q("p", "cNvPr"))
                shape_id = nv.attrib.get("id") if nv is not None else None
                if requested_shape is not None and shape_id != requested_shape:
                    continue
                if op_type == "setShapeText":
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
