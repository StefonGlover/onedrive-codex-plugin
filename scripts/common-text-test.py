#!/usr/bin/env python3
"""Dependency-free regression tests for common-text.py."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import zipfile
from email.message import EmailMessage
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
HELPER = ROOT / "scripts" / "common-text.py"


def extract(path: Path, kind: str) -> dict:
    request = json.dumps({"action": "extract", "inputPath": str(path), "kind": kind, "maxBytes": 16 * 1024})
    completed = subprocess.run(
        [sys.executable, str(HELPER)],
        input=request.encode("utf-8"),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        timeout=10,
    )
    result = json.loads(completed.stdout.decode("utf-8"))
    if completed.returncode != 0 or result.get("ok") is not True:
        raise AssertionError(result.get("error") or completed.stderr.decode("utf-8"))
    return result["value"]


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="onedrive-common-text-test-") as temporary:
        root = Path(temporary)

        rtf = root / "notes.rtf"
        rtf.write_text("{\\rtf1\\ansi Budget total: \\b $1,234\\b0\\par Next line}", encoding="utf-8")
        rtf_result = extract(rtf, "rtf")
        assert "Budget total: $1,234" in rtf_result["text"]
        assert "Next line" in rtf_result["text"]

        odt = root / "notes.odt"
        with zipfile.ZipFile(odt, "w") as package:
            package.writestr(
                "content.xml",
                '<?xml version="1.0"?><office:document-content '
                'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
                'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">'
                '<office:body><office:text><text:h>Budget</text:h><text:p>Income 12,325</text:p>'
                '</office:text></office:body></office:document-content>',
            )
        odt_result = extract(odt, "opendocument")
        assert "Budget" in odt_result["text"] and "Income 12,325" in odt_result["text"]

        epub = root / "book.epub"
        with zipfile.ZipFile(epub, "w") as package:
            package.writestr(
                "META-INF/container.xml",
                '<?xml version="1.0"?><container><rootfiles><rootfile full-path="OEBPS/content.opf"/></rootfiles></container>',
            )
            package.writestr(
                "OEBPS/content.opf",
                '<?xml version="1.0"?><package><manifest><item id="chapter" href="chapter.xhtml"/></manifest>'
                '<spine><itemref idref="chapter"/></spine></package>',
            )
            package.writestr("OEBPS/chapter.xhtml", "<html><body><h1>Chapter one</h1><p>Savings 3,350</p></body></html>")
        epub_result = extract(epub, "epub")
        assert "Chapter one" in epub_result["text"] and "Savings 3,350" in epub_result["text"]

        email = root / "service-report.eml"
        message = EmailMessage()
        message["Subject"] = "Evolution Heating service invoice 3095"
        message["From"] = "service@example.com"
        message["To"] = "owner@example.com"
        message["Date"] = "Wed, 24 Jul 2024 12:00:00 -0400"
        message.set_content("Compressor warranty replacement completed. Total paid: $2,518.")
        message.add_alternative("<html><body><p>HTML duplicate should not replace the plain body.</p></body></html>", subtype="html")
        message.add_attachment(b"not extracted", maintype="application", subtype="pdf", filename="invoice-3095.pdf")
        email.write_bytes(message.as_bytes())
        email_result = extract(email, "email")
        assert "Subject: Evolution Heating service invoice 3095" in email_result["text"]
        assert "Compressor warranty replacement completed" in email_result["text"]
        assert "Attachments: invoice-3095.pdf" in email_result["text"]
        assert "not extracted" not in email_result["text"]

    print(json.dumps({"ok": True, "tested": ["rtf", "opendocument", "epub", "email"]}, indent=2))


if __name__ == "__main__":
    main()
