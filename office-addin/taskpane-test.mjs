#!/usr/bin/env node
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { TextEncoder } from "node:util";

const elements = new Map(["host", "command", "apply", "result"].map((id) => [id, { id, disabled: true, value: "", textContent: "", addEventListener() {} }]));
let ready;
const requirements = new Set(["WordApi:1.3", "ExcelApi:1.7", "PowerPointApi:1.5"]);
const Office = {
  HostType: { Word: "Word", Excel: "Excel", PowerPoint: "PowerPoint" },
  context: { host: "Word", platform: "PC", requirements: { isSetSupported: (name, version) => requirements.has(`${name}:${version}`) } },
  onReady(callback) { ready = callback; }
};
const selection = { insertTextCalls: [], insertParagraphCalls: [], insertText(...args) { this.insertTextCalls.push(args); }, insertParagraph(...args) { this.insertParagraphCalls.push(args); } };
const Word = { InsertLocation: { replace: "Replace", before: "Before", after: "After" }, run: async (callback) => callback({ document: { getSelection: () => selection }, sync: async () => {} }) };
const range = { format: { fill: {}, font: {} }, clearCalls: [], clear(value) { this.clearCalls.push(value); } };
const Excel = { run: async (callback) => callback({ workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } }, sync: async () => {} }) };
const textRange = { text: "", font: {} };
const shapes = [{ textFrame: { hasText: true, textRange }, load() {}, deleted: false, delete() { this.deleted = true; } }];
const PowerPoint = { run: async (callback) => callback({ presentation: { getSelectedShapes: () => ({ items: shapes, load() {} }) }, sync: async () => {} }) };
const context = vm.createContext({ Office, Word, Excel, PowerPoint, TextEncoder, document: { getElementById: (id) => elements.get(id) }, console });
vm.runInContext(readFileSync(new URL("./taskpane.js", import.meta.url), "utf8"), context, { filename: "taskpane.js" });
await ready({ host: "Word", platform: "PC" });

const companion = context.CodexOfficeCompanion;
const wordCapabilities = companion.getCapabilities({ host: "Word", platform: "PC" });
assert.equal(wordCapabilities.protocolVersion, "codex-office-companion/1");
assert.deepEqual([...wordCapabilities.commands], ["replaceText", "setSelectedText", "insertParagraph"]);
assert.equal(wordCapabilities.transport.remoteCommands, false);
await companion.executeCommand({ protocolVersion: companion.protocolVersion, command: { type: "setSelectedText", text: "hello" } });
assert.deepEqual(selection.insertTextCalls[0], ["hello", "Replace"]);

Office.context.host = "Excel";
await companion.executeCommand({ type: "formatRange", sheet: "Data", address: "A1:B2", fillColor: "#ffeeaa", bold: true, fontSize: 14 });
assert.equal(range.format.fill.color, "#ffeeaa");
assert.equal(range.format.font.bold, true);
assert.equal(range.format.font.size, 14);
await companion.executeCommand({ type: "clearRange", sheet: "Data", address: "A1", applyTo: "All" });
assert.equal(range.clearCalls[0], "All");

Office.context.host = "PowerPoint";
await companion.executeCommand({ type: "setSelectedTextStyle", fontName: "Aptos", fontSize: 20, color: "#123456", bold: true });
assert.equal(textRange.font.name, "Aptos");
assert.equal(textRange.font.size, 20);
assert.equal(textRange.font.color, "#123456");
await companion.executeCommand({ type: "deleteSelectedShapes" });
assert.equal(shapes[0].deleted, true);

await assert.rejects(() => companion.executeCommand({ protocolVersion: "wrong", command: { type: "deleteSelectedShapes" } }), /protocolVersion/);
await assert.rejects(() => companion.executeCommand({ type: "deleteSelectedShapes", unknown: true }), /Unknown command property/);
console.log(JSON.stringify({ ok: true, hosts: ["Word", "Excel", "PowerPoint"], protocolVersion: companion.protocolVersion }));
