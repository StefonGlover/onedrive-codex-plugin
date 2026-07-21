#!/usr/bin/env python3
"""Bounded, dependency-light text extraction for common non-OpenXML files."""

from __future__ import annotations

import html
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Dict, Iterable, List
from xml.etree import ElementTree as ET


class ExtractionError(Exception):
    pass


def read_request() -> Dict[str, Any]:
    request = json.load(sys.stdin)
    if not isinstance(request, dict):
        raise ExtractionError("Request must be a JSON object.")
    return request


def bounded_text(payload: bytes, max_bytes: int) -> Dict[str, Any]:
    truncated = len(payload) > max_bytes
    text = payload[:max_bytes].decode("utf-8", errors="replace")
    if text.endswith("\ufffd"):
        text = text[:-1]
    return {"text": text.replace("\x00", ""), "truncated": truncated}


def safe_zip_member(package: zipfile.ZipFile, name: str, max_bytes: int = 20 * 1024 * 1024) -> bytes:
    try:
        info = package.getinfo(name)
    except KeyError as error:
        raise ExtractionError("Archive is missing required content: %s" % name) from error
    if info.file_size > max_bytes:
        raise ExtractionError("Archive member %s is too large to extract safely." % name)
    with package.open(info, "r") as source:
        payload = source.read(max_bytes + 1)
    if len(payload) > max_bytes:
        raise ExtractionError("Archive member %s exceeded the extraction limit." % name)
    return payload


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def structured_xml_text(payload: bytes) -> str:
    output: List[str] = []
    tab_elements = {"table-cell"}
    line_elements = {
        "h", "p", "list-item", "table", "table-row", "table-header-rows", "section",
        "draw-page", "notes", "title", "subtitle"
    }
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as error:
        raise ExtractionError("Document XML is malformed: %s" % error) from error

    def append_element(element: ET.Element) -> None:
        name = local_name(element.tag)
        if name == "tab":
            output.append("\t")
        elif name == "line-break":
            output.append("\n")
        if element.text:
            output.append(element.text)
        for child in list(element):
            append_element(child)
            if child.tail:
                output.append(child.tail)
        if name in tab_elements:
            output.append("\t")
        elif name in line_elements:
            output.append("\n")

    append_element(root)
    text = "".join(output)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return html.unescape(text).strip()


def extract_opendocument(path: Path) -> str:
    try:
        with zipfile.ZipFile(path, "r") as package:
            return structured_xml_text(safe_zip_member(package, "content.xml"))
    except zipfile.BadZipFile as error:
        raise ExtractionError("OpenDocument package is not a valid ZIP archive.") from error


class TextHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.output: List[str] = []
        self.ignored_depth = 0

    def handle_starttag(self, tag: str, attrs: List[Any]) -> None:
        if tag.lower() in {"script", "style"}:
            self.ignored_depth += 1
            return
        if self.ignored_depth:
            return
        if tag.lower() in {"br", "hr"}:
            self.output.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"script", "style"}:
            self.ignored_depth = max(0, self.ignored_depth - 1)
            return
        if self.ignored_depth:
            return
        if tag.lower() in {"address", "article", "aside", "blockquote", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "p", "pre", "section", "table", "td", "th", "tr"}:
            self.output.append("\n")

    def handle_data(self, data: str) -> None:
        if not self.ignored_depth:
            self.output.append(data)

    def text(self) -> str:
        value = "".join(self.output)
        value = re.sub(r"[ \t]+", " ", value)
        value = re.sub(r" *\n *", "\n", value)
        return re.sub(r"\n{3,}", "\n\n", value).strip()


def epub_content_members(package: zipfile.ZipFile) -> Iterable[str]:
    names = set(package.namelist())
    try:
        container = ET.fromstring(safe_zip_member(package, "META-INF/container.xml", 1024 * 1024))
        rootfile = next((node.attrib.get("full-path") for node in container.iter() if local_name(node.tag) == "rootfile"), None)
        if rootfile and rootfile in names:
            opf = ET.fromstring(safe_zip_member(package, rootfile, 5 * 1024 * 1024))
            manifest = {
                node.attrib.get("id"): node.attrib.get("href")
                for node in opf.iter() if local_name(node.tag) == "item"
            }
            base = Path(rootfile).parent
            for node in opf.iter():
                if local_name(node.tag) != "itemref":
                    continue
                href = manifest.get(node.attrib.get("idref"))
                if href:
                    name = str((base / href).as_posix())
                    if name in names:
                        yield name
            return
    except (ET.ParseError, ExtractionError):
        pass
    for name in sorted(names):
        if name.lower().endswith((".html", ".htm", ".xhtml")):
            yield name


def extract_epub(path: Path) -> str:
    try:
        with zipfile.ZipFile(path, "r") as package:
            chapters = []
            total = 0
            for name in epub_content_members(package):
                payload = safe_zip_member(package, name, 10 * 1024 * 1024)
                total += len(payload)
                if total > 50 * 1024 * 1024:
                    raise ExtractionError("EPUB expanded text exceeded the safety limit.")
                parser = TextHTMLParser()
                parser.feed(payload.decode("utf-8", errors="replace"))
                value = parser.text()
                if value:
                    chapters.append(value)
            if not chapters:
                raise ExtractionError("EPUB did not contain readable HTML chapters.")
            return "\n\n".join(chapters)
    except zipfile.BadZipFile as error:
        raise ExtractionError("EPUB is not a valid ZIP archive.") from error


def extract_email(path: Path) -> str:
    """Extract useful RFC 822 headers and bounded human-readable message bodies."""
    if path.stat().st_size > 25 * 1024 * 1024:
        raise ExtractionError("Email message is too large to parse safely.")
    try:
        with path.open("rb") as source:
            message = BytesParser(policy=policy.default).parse(source)
    except Exception as error:
        raise ExtractionError("Email message could not be parsed: %s" % error) from error

    output: List[str] = []
    for label, field in (("Subject", "subject"), ("From", "from"), ("To", "to"), ("Cc", "cc"), ("Date", "date")):
        value = message.get(field)
        if value:
            output.append("%s: %s" % (label, str(value).replace("\r", " ").replace("\n", " ")))

    plain_parts: List[str] = []
    html_parts: List[str] = []
    attachments: List[str] = []
    decoded_bytes = 0
    max_decoded_bytes = 8 * 1024 * 1024
    for part in message.walk():
        if part.is_multipart():
            continue
        filename = part.get_filename()
        disposition = (part.get_content_disposition() or "").lower()
        if filename or disposition == "attachment":
            if filename:
                attachments.append(str(filename).replace("\r", " ").replace("\n", " "))
            continue
        content_type = part.get_content_type().lower()
        if content_type not in {"text/plain", "text/html"}:
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            raw_payload = part.get_payload()
            payload = raw_payload.encode("utf-8", errors="replace") if isinstance(raw_payload, str) else b""
        remaining = max_decoded_bytes - decoded_bytes
        if remaining <= 0:
            break
        payload = payload[:remaining]
        decoded_bytes += len(payload)
        charset = part.get_content_charset() or "utf-8"
        try:
            value = payload.decode(charset, errors="replace")
        except LookupError:
            value = payload.decode("utf-8", errors="replace")
        if content_type == "text/html":
            parser = TextHTMLParser()
            parser.feed(value)
            value = parser.text()
            if value:
                html_parts.append(value)
        else:
            value = value.replace("\x00", "").strip()
            if value:
                plain_parts.append(value)

    bodies = plain_parts or html_parts
    if bodies:
        output.extend(["", *bodies])
    if attachments:
        output.extend(["", "Attachments: %s" % ", ".join(dict.fromkeys(attachments))])
    text = "\n".join(output).strip()
    if not text:
        raise ExtractionError("Email message did not contain readable headers or body text.")
    return text


RTF_DESTINATIONS = {
    "aftncn", "aftnsep", "aftnsepc", "annotation", "atnauthor", "atndate", "atnicn", "atnid",
    "atnparent", "atnref", "atntime", "atrfend", "atrfstart", "author", "background", "blipuid",
    "buptim", "category", "colorschememapping", "colortbl", "comment", "company", "creatim", "datafield",
    "datastore", "defchp", "defpap", "do", "doccomm", "docvar", "dptxbxtext", "ebcend", "ebcstart",
    "factoidname", "falt", "fchars", "ffdeftext", "ffentrymcr", "ffexitmcr", "ffformat", "ffhelptext",
    "ffl", "ffname", "ffstattext", "field", "file", "filetbl", "fldinst", "fldtype", "fname", "fontemb",
    "fontfile", "fonttbl", "footer", "footerf", "footerl", "footerr", "footnote", "formfield", "ftncn",
    "ftnsep", "ftnsepc", "generator", "header", "headerf", "headerl", "headerr", "hl", "hlfr", "hlinkbase",
    "htmltag", "info", "keycode", "keywords", "latentstyles", "lchars", "levelnumbers", "leveltext", "lfolevel",
    "linkval", "list", "listlevel", "listname", "listoverride", "listoverridetable", "listpicture", "liststylename",
    "listtable", "manager", "mhtmltag", "mmathPr", "nesttableprops", "nextfile", "nonesttables", "objalias",
    "objclass", "objdata", "object", "objname", "objsect", "objtime", "oldcprops", "oldpprops", "oldsprops",
    "oldtprops", "oleclsid", "operator", "panose", "password", "passwordhash", "pgp", "pgptbl", "picprop",
    "pict", "pn", "pnseclvl", "pntext", "pntxta", "pntxtb", "printim", "private", "propname", "protend",
    "protstart", "protusertbl", "pxe", "result", "revtbl", "revtim", "rsidtbl", "rxe", "shp", "shpgrp",
    "shpinst", "shppict", "shprslt", "shptxt", "sn", "sp", "staticval", "stylesheet", "subject", "sv",
    "svb", "tc", "template", "themedata", "title", "txe", "ud", "upr", "userprops", "wgrffmtfilter",
    "windowcaption", "writereservation", "writereservhash", "xe", "xmlattrname", "xmlattrvalue", "xmlclose",
    "xmlname", "xmlnstbl", "xmlopen"
}


def extract_rtf(path: Path) -> str:
    payload = path.read_bytes()
    if len(payload) > 20 * 1024 * 1024:
        raise ExtractionError("RTF is too large to parse safely.")
    source = payload.decode("latin-1", errors="replace")
    token = re.compile(r"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z0-9])|([{}])|\r\n|\r|\n|(.)", re.S)
    stack = []
    ignorable = False
    ucskip = 1
    curskip = 0
    output: List[str] = []
    for match in token.finditer(source):
        word, arg, hex_value, symbol, brace, char = match.groups()
        if brace:
            if brace == "{":
                stack.append((ignorable, ucskip))
            elif stack:
                ignorable, ucskip = stack.pop()
            continue
        if word:
            lower = word.lower()
            if lower in RTF_DESTINATIONS:
                ignorable = True
            elif lower == "uc" and arg:
                ucskip = max(0, int(arg))
            elif lower == "u" and arg:
                value = int(arg)
                if value < 0:
                    value += 65536
                if not ignorable:
                    output.append(chr(value))
                curskip = ucskip
            elif not ignorable and lower in {"par", "line"}:
                output.append("\n")
            elif not ignorable and lower == "tab":
                output.append("\t")
            continue
        if hex_value:
            if curskip:
                curskip -= 1
            elif not ignorable:
                output.append(bytes.fromhex(hex_value).decode("cp1252", errors="replace"))
            continue
        if symbol:
            if symbol == "*":
                ignorable = True
            elif not ignorable and symbol in "{}\\":
                output.append(symbol)
            elif not ignorable and symbol == "~":
                output.append(" ")
            continue
        if char:
            if curskip:
                curskip -= 1
            elif not ignorable:
                output.append(char)
    value = "".join(output)
    value = re.sub(r"[ \t]+\n", "\n", value)
    return re.sub(r"\n{3,}", "\n\n", value).strip()


EXTERNALS = {
    "pdf": (["/usr/bin/pdftotext", "/opt/homebrew/bin/pdftotext", "/usr/local/bin/pdftotext"], lambda exe, source, target: [exe, "-layout", "-nopgbrk", str(source), str(target)]),
    "image-ocr": (["/usr/bin/tesseract", "/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract"], lambda exe, source, target: [exe, str(source), str(target.with_suffix("")), "txt"]),
    "legacy-word": (["/usr/bin/catdoc", "/opt/homebrew/bin/catdoc", "/usr/local/bin/catdoc"], lambda exe, source, target: [exe, "-w", str(source)]),
    "legacy-excel": (["/usr/bin/xls2csv", "/opt/homebrew/bin/xls2csv", "/usr/local/bin/xls2csv"], lambda exe, source, target: [exe, str(source)]),
    "legacy-powerpoint": (["/usr/bin/catppt", "/opt/homebrew/bin/catppt", "/usr/local/bin/catppt"], lambda exe, source, target: [exe, str(source)]),
}


def extract_external(path: Path, kind: str, max_bytes: int) -> Dict[str, Any]:
    candidates, command = EXTERNALS[kind]
    executable = next((value for value in candidates if Path(value).is_file() and os.access(value, os.X_OK)), None)
    if not executable:
        raise ExtractionError("Required extractor is unavailable for %s." % kind)
    with tempfile.TemporaryDirectory(prefix="onedrive-extract-", dir=str(path.parent)) as temporary:
        output = Path(temporary) / "output.txt"
        args = command(executable, path, output)
        writes_output_file = kind in {"pdf", "image-ocr"}
        try:
            if writes_output_file:
                completed = subprocess.run(args, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, timeout=55, check=False)
            else:
                with output.open("wb") as destination:
                    completed = subprocess.run(args, stdout=destination, stderr=subprocess.PIPE, timeout=55, check=False)
        except subprocess.TimeoutExpired as error:
            raise ExtractionError("%s extraction timed out." % kind) from error
        if completed.returncode != 0:
            message = completed.stderr[:4096].decode("utf-8", errors="replace").strip()
            raise ExtractionError("%s extraction failed: %s" % (kind, message or "unknown extractor error"))
        if not output.is_file():
            raise ExtractionError("%s extractor did not produce text." % kind)
        with output.open("rb") as source:
            payload = source.read(max_bytes + 1)
        result = bounded_text(payload, max_bytes)
        result["extractor"] = executable
        return result


def extract(path: Path, kind: str, max_bytes: int) -> Dict[str, Any]:
    if kind == "email":
        result = bounded_text(extract_email(path).encode("utf-8"), max_bytes)
        result["extractor"] = "python-email"
        return result
    if kind == "rtf":
        result = bounded_text(extract_rtf(path).encode("utf-8"), max_bytes)
        result["extractor"] = "python-rtf"
        return result
    if kind == "opendocument":
        result = bounded_text(extract_opendocument(path).encode("utf-8"), max_bytes)
        result["extractor"] = "python-opendocument"
        return result
    if kind == "epub":
        result = bounded_text(extract_epub(path).encode("utf-8"), max_bytes)
        result["extractor"] = "python-epub"
        return result
    if kind in EXTERNALS:
        return extract_external(path, kind, max_bytes)
    raise ExtractionError("Unsupported extraction kind: %s" % kind)


def main() -> None:
    try:
        request = read_request()
        if request.get("action") != "extract":
            raise ExtractionError("action must be extract.")
        input_path = request.get("inputPath")
        if not isinstance(input_path, str) or not input_path:
            raise ExtractionError("inputPath is required.")
        path = Path(input_path).expanduser().resolve()
        if not path.is_file():
            raise ExtractionError("Input file does not exist.")
        kind = str(request.get("kind") or "")
        max_bytes = max(1, min(int(request.get("maxBytes", 192 * 1024)), 1024 * 1024))
        value = extract(path, kind, max_bytes)
        print(json.dumps({"ok": True, "value": value}, ensure_ascii=False, separators=(",", ":")))
    except Exception as error:
        print(json.dumps({"ok": False, "error": str(error)}, ensure_ascii=False, separators=(",", ":")))
        sys.exit(1)


if __name__ == "__main__":
    main()
