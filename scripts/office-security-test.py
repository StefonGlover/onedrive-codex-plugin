#!/usr/bin/env python3
"""Adversarial package and split-run regression tests for the Office helper."""

import json
import os
import random
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
HELPER = ROOT / "scripts" / "office-openxml.py"
FIXTURES = ROOT / "scripts" / "office-openxml-test.py"


def call(path, action="validate", **options):
    result = subprocess.run(
        [sys.executable, str(HELPER)],
        input=json.dumps({"action": action, "inputPath": str(path), **options}),
        text=True,
        capture_output=True,
        check=False,
        env={**os.environ, "PYTHONPYCACHEPREFIX": str(path.parent / "pycache")},
        timeout=15,
    )
    payload = json.loads(result.stdout or "{}")
    return result.returncode, payload


def rewrite(source, destination, replacements=None, additions=None, omit=None):
    replacements = replacements or {}
    additions = additions or {}
    omit = set(omit or [])
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as outgoing:
        for info in incoming.infolist():
            if info.filename in omit:
                continue
            outgoing.writestr(info, replacements.get(info.filename, incoming.read(info.filename)))
        for name, value in additions.items():
            outgoing.writestr(name, value)


def main():
    checks = {}
    with tempfile.TemporaryDirectory(prefix="onedrive-office-security-") as directory:
        root = Path(directory)
        fixtures = root / "fixtures"
        subprocess.run([sys.executable, str(FIXTURES), f"--emit-fixtures={fixtures}"], check=True, capture_output=True, text=True)
        source = fixtures / "sample.docx"

        duplicate = root / "duplicate.docx"
        rewrite(source, duplicate)
        with zipfile.ZipFile(duplicate, "a") as package:
            package.writestr("word/document.xml", b"duplicate")
        code, payload = call(duplicate, kind="word")
        checks["duplicateMembersRejected"] = code != 0 and "duplicate ZIP member" in payload.get("error", "")

        unsafe = root / "unsafe.docx"
        rewrite(source, unsafe, additions={"../escape.xml": b"<x/>"})
        code, payload = call(unsafe, kind="word")
        checks["unsafePathsRejected"] = code != 0 and "unsafe ZIP paths" in payload.get("error", "")

        missing_types = root / "missing-types.docx"
        rewrite(source, missing_types, omit={"[Content_Types].xml"})
        code, payload = call(missing_types, kind="word")
        checks["missingContentTypesRejected"] = code != 0 and "content types" in payload.get("error", "").lower()

        malformed = root / "malformed.docx"
        rewrite(source, malformed, replacements={"word/document.xml": b"<w:document>"})
        code, payload = call(malformed, kind="word")
        checks["malformedXmlRejected"] = code != 0 and "Invalid XML" in payload.get("error", "")

        bomb = root / "bomb.docx"
        rewrite(source, bomb, additions={"word/huge.xml": b"A" * (2 * 1024 * 1024)})
        code, payload = call(bomb, kind="word")
        checks["compressionBombRejected"] = code != 0 and "suspiciously compressed" in payload.get("error", "")

        split = root / "split.docx"
        split_xml = b'''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello </w:t></w:r><w:r><w:t>split run</w:t></w:r></w:p></w:body></w:document>'''
        rewrite(source, split, replacements={"word/document.xml": split_xml})
        edited = root / "split-edited.docx"
        code, payload = call(split, action="edit", kind="word", outputPath=str(edited), operations=[{"type": "replaceText", "find": "Hello split", "replace": "Safe cross-run"}])
        with zipfile.ZipFile(edited) as package:
            edited_xml = package.read("word/document.xml")
        checks["splitRunReplacement"] = code == 0 and payload.get("value", {}).get("changeCount") == 1 and b"Safe cross-run" in edited_xml and b"<w:b" in edited_xml

        with zipfile.ZipFile(source) as before, zipfile.ZipFile(edited) as after:
            untouched = [name for name in before.namelist() if name != "word/document.xml"]
            checks["untouchedPartsPreserved"] = all(before.read(name) == after.read(name) for name in untouched)

        fuzz_ok = True
        randomizer = random.Random(20260711)
        for index in range(20):
            payload_bytes = bytearray(split_xml)
            for _ in range(1 + index % 5):
                position = randomizer.randrange(len(payload_bytes))
                payload_bytes[position] = randomizer.randrange(256)
            candidate = root / ("fuzz-%02d.docx" % index)
            rewrite(source, candidate, replacements={"word/document.xml": bytes(payload_bytes)})
            try:
                _, response = call(candidate, kind="word")
                fuzz_ok = fuzz_ok and isinstance(response.get("ok"), bool)
            except Exception:
                fuzz_ok = False
                break
        checks["boundedMalformedFuzz"] = fuzz_ok

        ppt_source = fixtures / "sample.pptx"
        with zipfile.ZipFile(ppt_source) as package:
            base_slide = package.read("ppt/slides/slide1.xml")
        split_property_ok = True
        phrase = "Hello PowerPoint"
        for index in range(1, len(phrase)):
            runs = ("<a:r><a:rPr lang=\"en-US\"/><a:t>%s</a:t></a:r><a:r><a:rPr b=\"1\"/><a:t>%s</a:t></a:r>" % (phrase[:index], phrase[index:])).encode()
            split_slide = re.sub(rb"<a:r><a:t>Hello PowerPoint</a:t></a:r>", runs, base_slide)
            candidate = root / ("split-ppt-%02d.pptx" % index)
            output = root / ("split-ppt-%02d-edited.pptx" % index)
            rewrite(ppt_source, candidate, replacements={"ppt/slides/slide1.xml": split_slide})
            code, response = call(candidate, action="edit", kind="powerpoint", outputPath=str(output), operations=[{"type": "replaceText", "slideIndex": 0, "shapeId": "2", "find": phrase, "replace": "Cross-run safe"}])
            if code != 0 or response.get("value", {}).get("changeCount") != 1:
                split_property_ok = False
                break
            with zipfile.ZipFile(output) as package:
                edited_slide = package.read("ppt/slides/slide1.xml")
            if b"Cross-run safe" not in edited_slide or b'b="1"' not in edited_slide:
                split_property_ok = False
                break
        checks["powerpointSplitRunProperty"] = split_property_ok

        large = root / "large-runs.pptx"
        large_edited = root / "large-runs-edited.pptx"
        large_runs = b"".join((b"<a:r><a:t>chunk%04d </a:t></a:r>" % index) for index in range(5000)) + b"<a:r><a:t>END-MARKER</a:t></a:r>"
        large_slide = re.sub(rb"<a:r><a:t>Hello PowerPoint</a:t></a:r>", large_runs, base_slide)
        rewrite(ppt_source, large, replacements={"ppt/slides/slide1.xml": large_slide})
        code, response = call(large, action="edit", kind="powerpoint", outputPath=str(large_edited), operations=[{"type": "replaceText", "slideIndex": 0, "shapeId": "2", "find": "END-MARKER", "replace": "LARGE-OK"}])
        checks["largeRunCorpus"] = code == 0 and response.get("value", {}).get("changeCount") == 1 and large_edited.stat().st_size > 0

        traversal_rel = root / "relationship-traversal.pptx"
        bad_rels = b'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../../../outside.xml"/></Relationships>'''
        rewrite(ppt_source, traversal_rel, replacements={"ppt/_rels/presentation.xml.rels": bad_rels})
        code, response = call(traversal_rel, kind="powerpoint")
        checks["relationshipTraversalRejected"] = code != 0 and "broken internal relationships" in response.get("error", "")

        truncated = root / "truncated.pptx"
        truncated.write_bytes(ppt_source.read_bytes()[: max(1, ppt_source.stat().st_size // 2)])
        code, response = call(truncated, kind="powerpoint")
        checks["truncatedZipRejected"] = code != 0 and response.get("ok") is False

    print(json.dumps({"ok": all(checks.values()), "checks": checks}, indent=2))
    raise SystemExit(0 if all(checks.values()) else 1)


if __name__ == "__main__":
    main()
