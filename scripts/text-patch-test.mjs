#!/usr/bin/env node
import assert from "node:assert/strict";
import { applyTextPatch, decodeTextBuffer, encodeTextBuffer } from "../mcp/text-patch.mjs";

let passed = 0;
const check = (name, callback) => { callback(); passed += 1; process.stdout.write(`ok ${passed} - ${name}\n`); };

check("unified patches preserve UTF-8 BOM, CRLF, and trailing newline", () => {
  const source = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("one\r\ntwo\r\n")]);
  const result = applyTextPatch(source, { mode: "unified", diff: "--- a\n+++ b\n@@ -1,2 +1,2 @@\n one\n-two\n+changed" });
  assert.deepEqual([...result.bytes.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
  assert.equal(decodeTextBuffer(result.bytes).text, "one\r\nchanged\r\n");
});

check("JSON Patch supports escaped pointers and UTF-16LE preservation", () => {
  const metadata = { encoding: "utf16le", bom: Buffer.from([0xff, 0xfe]), newline: "\n", trailingNewline: false };
  const source = encodeTextBuffer('{"a/b":{"~key":1}}', metadata);
  const result = applyTextPatch(source, { mode: "json", operations: [{ op: "replace", path: "/a~1b/~0key", value: 2 }] });
  const decoded = decodeTextBuffer(result.bytes);
  assert.equal(decoded.encoding, "utf16le");
  assert.deepEqual(JSON.parse(decoded.text), { "a/b": { "~key": 2 } });
});

check("UTF-16BE content retains byte order and BOM", () => {
  const metadata = { encoding: "utf16be", bom: Buffer.from([0xfe, 0xff]), newline: "\n", trailingNewline: true };
  const source = encodeTextBuffer('{"value":1}\n', metadata);
  const result = applyTextPatch(source, { mode: "json", operations: [{ op: "replace", path: "/value", value: 3 }] });
  assert.deepEqual([...result.bytes.subarray(0, 2)], [0xfe, 0xff]);
  assert.equal(JSON.parse(decodeTextBuffer(result.bytes).text).value, 3);
});

check("safe YAML rejects aliases, custom tags, merge keys, and duplicate keys", () => {
  for (const source of ["a: &x 1\nb: *x\n", "a: !tag value\n", "a: 1\n<<: x\n", "a: 1\na: 2\n"]) {
    assert.throws(() => applyTextPatch(Buffer.from(source), { mode: "yaml", operations: [{ op: "replace", path: "/a", value: 2 }] }));
  }
});

check("CSV row-key updates preserve RFC 4180 quoting", () => {
  const source = Buffer.from('id,name,notes\r\n1,"Doe, Jane","said ""hi"""\r\n');
  const result = applyTextPatch(source, { mode: "csv", keyColumn: "id", operations: [{ op: "update", key: "1", values: [{ column: "notes", value: "line 1\nline 2" }] }, { op: "insert", key: "2", values: [{ column: "name", value: "Smith" }, { column: "notes", value: "ok" }] }] });
  const text = decodeTextBuffer(result.bytes).text;
  assert.match(text, /"Doe, Jane"/);
  assert.match(text, /"line 1\r\nline 2"/);
  assert.match(text, /2,Smith,ok/);
});

check("malformed patches and binary input fail closed", () => {
  assert.throws(() => applyTextPatch(Buffer.from("one\n"), { mode: "unified", diff: "not a diff" }), /no hunk/);
  assert.throws(() => applyTextPatch(Buffer.from([0, 1, 2, 3]), { mode: "json", operations: [] }), /NUL bytes/);
  assert.throws(() => applyTextPatch(Buffer.from("{}"), { mode: "json", operations: [{ op: "remove", path: "/missing" }] }), /does not exist/);
});

console.log(JSON.stringify({ ok: true, passed }));
