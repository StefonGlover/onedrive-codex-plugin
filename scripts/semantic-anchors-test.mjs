#!/usr/bin/env node
import assert from "node:assert/strict";
import { addSemanticAnchors, resolveSemanticOperations } from "../mcp/semantic-anchors.mjs";

let passed = 0;
const check = (name, callback) => { callback(); passed += 1; process.stdout.write(`ok ${passed} - ${name}\n`); };

check("Word anchors rebase a uniquely moved paragraph", () => {
  const before = addSemanticAnchors("word", { paragraphs: [
    { index: 0, part: "word/document.xml", style: "Heading 1", text: "Scope" },
    { index: 1, part: "word/document.xml", style: "Normal", text: "Durable target" },
    { index: 2, part: "word/document.xml", style: "Normal", text: "Tail" }
  ] });
  const anchor = before.paragraphs[1].anchor;
  const after = addSemanticAnchors("word", { paragraphs: [
    { index: 0, part: "word/document.xml", style: "Normal", text: "New lead" },
    { index: 1, part: "word/document.xml", style: "Heading 1", text: "Scope" },
    { index: 2, part: "word/document.xml", style: "Normal", text: "Durable target" },
    { index: 3, part: "word/document.xml", style: "Normal", text: "Tail" }
  ] });
  const result = resolveSemanticOperations("word", after, [{ type: "setParagraphText", anchor, text: "Changed" }]);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.operations[0].paragraphIndex, 2);
});

check("duplicate and missing Word anchors fail closed", () => {
  const inspection = addSemanticAnchors("word", { paragraphs: [
    { index: 0, part: "word/document.xml", text: "same" },
    { index: 1, part: "word/document.xml", text: "same" }
  ] });
  const duplicateAnchor = { ...inspection.paragraphs[0].anchor, fingerprint: undefined, beforeHash: null, afterHash: null };
  const duplicate = resolveSemanticOperations("word", inspection, [{ type: "setParagraphText", anchor: duplicateAnchor, text: "x" }]);
  assert.equal(duplicate.conflicts[0].reason, "ambiguous_anchor");
  const missing = resolveSemanticOperations("word", inspection, [{ type: "setParagraphText", anchor: { ...duplicateAnchor, textHash: "missing" }, text: "x" }]);
  assert.equal(missing.conflicts[0].reason, "anchor_not_found");
});

check("anchor and legacy selector must resolve to the same Word object", () => {
  const inspection = addSemanticAnchors("word", { paragraphs: [{ index: 0, part: "word/document.xml", text: "A" }, { index: 1, part: "word/document.xml", text: "B" }] });
  const result = resolveSemanticOperations("word", inspection, [{ type: "setParagraphText", paragraphIndex: 0, anchor: inspection.paragraphs[1].anchor, text: "x" }]);
  assert.equal(result.conflicts[0].reason, "selector_anchor_mismatch");
});

check("Excel anchors emit on helper sheets and follow uniquely moved cells", () => {
  const before = addSemanticAnchors("excel", { sheets: [{ name: "Data", relationshipId: "rId1", part: "xl/worksheets/sheet1.xml", cells: [{ address: "A1", value: "key", formula: null }] }] });
  assert.equal(before.sheets[0].cells[0].anchor.kind, "range");
  const after = addSemanticAnchors("excel", { sheets: [{ name: "Data", relationshipId: "rId1", part: "xl/worksheets/sheet1.xml", cells: [{ address: "C4", value: "key", formula: null }] }] });
  const result = resolveSemanticOperations("excel", after, [{ type: "setCell", anchor: before.sheets[0].cells[0].anchor, value: "next" }]);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.operations[0].address, "C4");
});

check("PowerPoint persistent slide and shape IDs survive positional movement", () => {
  const before = addSemanticAnchors("powerpoint", { slides: [{ index: 0, relationshipId: "rId7", shapes: [{ id: 4, name: "Body", text: "Target" }] }] });
  const after = addSemanticAnchors("powerpoint", { slides: [{ index: 3, relationshipId: "rId7", shapes: [{ id: 4, name: "Body", text: "Target" }] }] });
  const result = resolveSemanticOperations("powerpoint", after, [{ type: "setShapeText", anchor: before.slides[0].shapes[0].anchor, text: "Changed" }]);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.operations[0].slideIndex, 3);
  assert.equal(result.operations[0].shapeId, 4);
});

check("rebasePolicy fail requires an exact unchanged fingerprint", () => {
  const before = addSemanticAnchors("word", { paragraphs: [{ index: 0, part: "word/document.xml", text: "A" }, { index: 1, part: "word/document.xml", text: "Target" }] });
  const moved = addSemanticAnchors("word", { paragraphs: [{ index: 0, part: "word/document.xml", text: "Target" }, { index: 1, part: "word/document.xml", text: "New" }, { index: 2, part: "word/document.xml", text: "A" }] });
  const result = resolveSemanticOperations("word", moved, [{ type: "setParagraphText", anchor: before.paragraphs[1].anchor, rebasePolicy: "fail", text: "x" }]);
  assert.equal(result.conflicts.length, 1);
});

console.log(JSON.stringify({ ok: true, passed }));
