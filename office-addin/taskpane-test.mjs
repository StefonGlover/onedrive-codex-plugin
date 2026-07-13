#!/usr/bin/env node
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { TextEncoder } from "node:util";
import { OFFICE_ICON_SIZES, officeManifestProblems } from "./manifest-contract.mjs";

let passed = 0;
async function check(name, callback) {
  await callback();
  passed += 1;
  process.stdout.write(`ok ${passed} - ${name}\n`);
}

const elements = new Map(["host", "command", "apply", "result"].map((id) => [id, {
  id,
  disabled: true,
  value: "",
  textContent: "",
  listeners: {},
  addEventListener(name, callback) { this.listeners[name] = callback; }
}]));
let ready;
const requirements = new Set(["WordApi:1.3", "ExcelApi:1.7", "PowerPointApi:1.5"]);
const Office = {
  HostType: { Word: "Word", Excel: "Excel", PowerPoint: "PowerPoint" },
  context: { host: "Word", platform: "PC", requirements: { isSetSupported: (name, version) => requirements.has(`${name}:${version}`) } },
  onReady(callback) { ready = callback; }
};

const wordSearchRanges = [
  { inserted: [], insertText(...args) { this.inserted.push(args); } },
  { inserted: [], insertText(...args) { this.inserted.push(args); } }
];
let wordSearchRequest;
const wordSelection = {
  isEmpty: false,
  insertTextCalls: [],
  insertParagraphCalls: [],
  load() {},
  insertText(...args) { this.insertTextCalls.push(args); },
  insertParagraph(...args) { this.insertParagraphCalls.push(args); }
};
const Word = {
  InsertLocation: { replace: "Replace", before: "Before", after: "After" },
  run: async (callback) => callback({
    document: {
      body: {
        search(find, options) {
          wordSearchRequest = { find, options };
          return { items: wordSearchRanges, load() {} };
        }
      },
      getSelection: () => wordSelection
    },
    sync: async () => {}
  })
};

const excelRanges = [];
const rangeDimensions = new Map([
  ["A1", { rowCount: 1, columnCount: 1 }],
  ["A1:B2", { rowCount: 2, columnCount: 2 }],
  ["C3:D3", { rowCount: 1, columnCount: 2 }]
]);
function makeExcelRange(address) {
  const dimensions = rangeDimensions.get(address) || { rowCount: 1, columnCount: 1 };
  const range = {
    ...dimensions,
    address,
    format: { fill: {}, font: {} },
    clearCalls: [],
    load() {},
    clear(value) { this.clearCalls.push(value); }
  };
  excelRanges.push(range);
  return range;
}
const Excel = {
  run: async (callback) => callback({
    workbook: { worksheets: { getItem: (sheet) => ({ getRange: (address) => Object.assign(makeExcelRange(address), { sheet }) }) } },
    sync: async () => {}
  })
};

const selectedTextRange = { isNullObject: false, text: "selected", font: {}, load() {} };
let selectedShapes = [
  { deleted: false, delete() { this.deleted = true; } },
  { deleted: false, delete() { this.deleted = true; } }
];
let selectedTextRangeCalls = 0;
let selectedShapeCalls = 0;
const PowerPoint = {
  ShapeFontUnderlineStyle: { none: "None", single: "Single" },
  run: async (callback) => callback({
    presentation: {
      getSelectedTextRangeOrNullObject() {
        selectedTextRangeCalls += 1;
        return selectedTextRange;
      },
      getSelectedShapes() {
        selectedShapeCalls += 1;
        return { items: selectedShapes, load() {} };
      }
    },
    sync: async () => {}
  })
};

const context = vm.createContext({
  Office,
  Word,
  Excel,
  PowerPoint,
  TextEncoder,
  document: { getElementById: (id) => elements.get(id) },
  console
});
vm.runInContext(readFileSync(new URL("./taskpane.js", import.meta.url), "utf8"), context, { filename: "taskpane.js" });
await ready({ host: "Word", platform: "PC" });
const companion = context.CodexOfficeCompanion;

async function rejectFor(host, command, pattern) {
  Office.context.host = host;
  await assert.rejects(() => companion.executeCommand(command), pattern);
}

await check("capabilities preserve protocol and advertise companion 1.1.1", async () => {
  const expected = {
    Word: ["replaceText", "setSelectedText", "insertParagraph"],
    Excel: ["setRange", "clearRange", "formatRange"],
    PowerPoint: ["setSelectedText", "setSelectedTextStyle", "deleteSelectedShapes"]
  };
  for (const [host, commands] of Object.entries(expected)) {
    const capabilities = companion.getCapabilities({ host, platform: "PC" });
    assert.equal(capabilities.protocolVersion, "codex-office-companion/1");
    assert.equal(capabilities.companionVersion, "1.1.1");
    assert.deepEqual([...capabilities.commands], commands);
    assert.equal(capabilities.limits.maxCommandBytes, 65536);
    assert.equal(capabilities.transport.remoteCommands, false);
    assert.equal(capabilities.transport.telemetry, false);
  }
});

await check("host-specific starter commands are valid and replace the textarea placeholder", async () => {
  for (const host of ["Word", "Excel", "PowerPoint"]) {
    Office.context.host = host;
    await ready({ host, platform: "PC" });
    const starter = JSON.parse(elements.get("command").value);
    assert.ok(companion.getCapabilities({ host }).commands.includes(starter.type));
  }
});

await check("Word replaceText honors envelope negotiation, match options, and first-only mode", async () => {
  Office.context.host = "Word";
  const response = await companion.executeCommand({
    protocolVersion: companion.protocolVersion,
    command: { type: "replaceText", find: "Old", replace: "New", all: false, matchCase: true }
  });
  assert.equal(response.operation, "replaceText");
  assert.equal(response.changed, 1);
  assert.equal(wordSearchRequest.find, "Old");
  assert.equal(wordSearchRequest.options.matchCase, true);
  assert.equal(wordSearchRequest.options.matchWholeWord, false);
  assert.deepEqual(wordSearchRanges[0].inserted[0], ["New", "Replace"]);
  assert.equal(wordSearchRanges[1].inserted.length, 0);
});

await check("Word setSelectedText replaces a non-empty selection", async () => {
  Office.context.host = "Word";
  await companion.executeCommand({ type: "setSelectedText", text: "hello" });
  assert.deepEqual(wordSelection.insertTextCalls.at(-1), ["hello", "Replace"]);
});

await check("Word insertParagraph validates and defaults location to after", async () => {
  Office.context.host = "Word";
  await companion.executeCommand({ type: "insertParagraph", text: "next paragraph" });
  assert.deepEqual(wordSelection.insertParagraphCalls.at(-1), ["next paragraph", "After"]);
  await companion.executeCommand({ type: "insertParagraph", text: "previous paragraph", location: "before" });
  assert.deepEqual(wordSelection.insertParagraphCalls.at(-1), ["previous paragraph", "Before"]);
});

await check("Excel setRange accepts rectangular matrices matching the target range", async () => {
  Office.context.host = "Excel";
  await companion.executeCommand({
    type: "setRange",
    sheet: "Data",
    address: "A1:B2",
    values: [[1, 2], [3, 4]],
    formulas: [["=1", "=2"], ["=3", "=4"]],
    numberFormat: [["0", "0"], ["0", "0"]]
  });
  const range = excelRanges.at(-1);
  assert.deepEqual(range.values, [[1, 2], [3, 4]]);
  assert.deepEqual(range.formulas, [["=1", "=2"], ["=3", "=4"]]);
  assert.deepEqual(range.numberFormat, [["0", "0"], ["0", "0"]]);
});

await check("Excel clearRange supports every ExcelApi 1.7 clear mode and defaults to Contents", async () => {
  Office.context.host = "Excel";
  await companion.executeCommand({ type: "clearRange", sheet: "Data", address: "A1" });
  assert.equal(excelRanges.at(-1).clearCalls[0], "Contents");
  for (const applyTo of ["All", "Contents", "Formats", "Hyperlinks", "RemoveHyperlinks"]) {
    await companion.executeCommand({ type: "clearRange", sheet: "Data", address: "A1", applyTo });
    assert.equal(excelRanges.at(-1).clearCalls[0], applyTo);
  }
});

await check("Excel formatRange applies typed font, fill, and number-format values", async () => {
  Office.context.host = "Excel";
  await companion.executeCommand({
    type: "formatRange",
    sheet: "Data",
    address: "C3:D3",
    fillColor: "#ffeeaa",
    fontColor: "#112233",
    fontName: "Aptos",
    fontSize: 14,
    bold: true,
    italic: false,
    numberFormat: [["0.00", "0.00"]]
  });
  const range = excelRanges.at(-1);
  assert.equal(range.format.fill.color, "#ffeeaa");
  assert.deepEqual(range.format.font, { color: "#112233", name: "Aptos", size: 14, bold: true, italic: false });
  assert.deepEqual(range.numberFormat, [["0.00", "0.00"]]);
});

await check("PowerPoint setSelectedText uses only the selected text range", async () => {
  Office.context.host = "PowerPoint";
  selectedTextRange.isNullObject = false;
  const shapeCallsBefore = selectedShapeCalls;
  await companion.executeCommand({ type: "setSelectedText", text: "replacement" });
  assert.equal(selectedTextRange.text, "replacement");
  assert.equal(selectedShapeCalls, shapeCallsBefore);
  assert.ok(selectedTextRangeCalls > 0);
});

await check("PowerPoint setSelectedTextStyle maps underline booleans to Office enums", async () => {
  Office.context.host = "PowerPoint";
  await companion.executeCommand({ type: "setSelectedTextStyle", fontName: "Aptos", fontSize: 20, color: "#123456", bold: true, italic: false, underline: true });
  assert.deepEqual(selectedTextRange.font, { name: "Aptos", size: 20, color: "#123456", bold: true, italic: false, underline: "Single" });
  await companion.executeCommand({ type: "setSelectedTextStyle", underline: false });
  assert.equal(selectedTextRange.font.underline, "None");
});

await check("PowerPoint deleteSelectedShapes uses only the selected-shape collection", async () => {
  Office.context.host = "PowerPoint";
  const textCallsBefore = selectedTextRangeCalls;
  const response = await companion.executeCommand({ type: "deleteSelectedShapes" });
  assert.equal(response.changed, 2);
  assert.ok(selectedShapes.every((shape) => shape.deleted));
  assert.equal(selectedTextRangeCalls, textCallsBefore);
});

await check("executeCommand enforces the 64 KiB UTF-8 serialized-command limit", async () => {
  await rejectFor("Word", { type: "setSelectedText", text: "😀".repeat(20_000) }, /exceeds 65536 bytes/);
  const cyclic = { type: "setSelectedText", text: "hello" };
  cyclic.self = cyclic;
  await rejectFor("Word", cyclic, /JSON-serializable/);
});

await check("envelopes and command objects reject malformed protocol shapes", async () => {
  await rejectFor("PowerPoint", { protocolVersion: "wrong", command: { type: "deleteSelectedShapes" } }, /protocolVersion/);
  await rejectFor("PowerPoint", { command: { type: "deleteSelectedShapes" } }, /protocolVersion/);
  await rejectFor("PowerPoint", [], /JSON object/);
  await rejectFor("PowerPoint", { type: 3 }, /type must be a string/);
  await rejectFor("PowerPoint", { type: "deleteSelectedShapes", unknown: true }, /Unknown command property/);
});

await check("Word commands reject invalid types, enums, and empty selection", async () => {
  await rejectFor("Word", { type: "replaceText", find: "" }, /find must not be empty/);
  await rejectFor("Word", { type: "replaceText", find: "x", replace: 1 }, /replace must be a string/);
  await rejectFor("Word", { type: "replaceText", find: "x", all: "yes" }, /all must be a boolean/);
  await rejectFor("Word", { type: "insertParagraph", text: "x", location: "start" }, /location must be one of/);
  await rejectFor("Word", { type: "setSelectedText", text: null }, /text must be a string/);
  wordSelection.isEmpty = true;
  await rejectFor("Word", { type: "setSelectedText", text: "x" }, /non-empty Word text selection/);
  wordSelection.isEmpty = false;
});

await check("Excel commands reject no-ops, invalid matrices, bad dimensions, and non-finite numbers", async () => {
  await rejectFor("Excel", { type: "setRange", sheet: "Data", address: "A1" }, /requires values/);
  await rejectFor("Excel", { type: "formatRange", sheet: "Data", address: "A1" }, /at least one formatting property/);
  await rejectFor("Excel", { type: "setRange", sheet: "Data", address: "A1:B2", values: [[1], [2, 3]] }, /rectangular/);
  await rejectFor("Excel", { type: "setRange", sheet: "Data", address: "A1:B2", values: [[1]] }, /must match range dimensions/);
  await rejectFor("Excel", { type: "setRange", sheet: "Data", address: "A1", values: [[NaN]] }, /finite number/);
  await rejectFor("Excel", { type: "clearRange", sheet: "Data", address: "A1", applyTo: "ResetContents" }, /applyTo must be one of/);
  await rejectFor("Excel", { type: "formatRange", sheet: "Data", address: "A1", bold: 1 }, /bold must be a boolean/);
  await rejectFor("Excel", { type: "formatRange", sheet: "Data", address: "A1", fontSize: Infinity }, /finite number/);
  await rejectFor("Excel", { type: "formatRange", sheet: "Data", address: "A1", numberFormat: [[1]] }, /must be a string/);
});

await check("PowerPoint commands reject no-ops, invalid styles, and missing selections", async () => {
  await rejectFor("PowerPoint", { type: "setSelectedTextStyle" }, /at least one formatting property/);
  await rejectFor("PowerPoint", { type: "setSelectedTextStyle", underline: "Single" }, /underline must be a boolean/);
  await rejectFor("PowerPoint", { type: "setSelectedTextStyle", fontSize: 0 }, /between 1 and 400/);
  await rejectFor("PowerPoint", { type: "setSelectedText", text: 9 }, /text must be a string/);
  selectedTextRange.isNullObject = true;
  await rejectFor("PowerPoint", { type: "setSelectedTextStyle", bold: true }, /non-empty PowerPoint text selection/);
  selectedTextRange.isNullObject = false;
  const previousShapes = selectedShapes;
  selectedShapes = [];
  await rejectFor("PowerPoint", { type: "deleteSelectedShapes" }, /at least one selected shape/);
  selectedShapes = previousShapes;
});

await check("unavailable requirement sets fail closed", async () => {
  Office.context.host = "Word";
  requirements.delete("WordApi:1.3");
  assert.equal(companion.getCapabilities({ host: "Word" }).available, false);
  await assert.rejects(() => companion.executeCommand({ type: "setSelectedText", text: "x" }), /does not satisfy/);
  requirements.add("WordApi:1.3");
});

await check("manifest exposes one validated ShowTaskpane ribbon command in every supported host", async () => {
  const manifest = readFileSync(new URL("./manifest.xml", import.meta.url), "utf8");
  assert.deepEqual(officeManifestProblems(manifest), []);
  assert.equal((manifest.match(/xsi:type="ShowTaskpane"/g) || []).length, 3);
  assert.doesNotMatch(manifest, /ExecuteFunction/);
});

await check("base and ribbon Office icons have their exact advertised PNG dimensions", async () => {
  for (const size of OFFICE_ICON_SIZES) {
    const name = `icon-${size}.png`;
    const png = readFileSync(new URL(`./${name}`, import.meta.url));
    assert.deepEqual([...png.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
    assert.equal(png.readUInt32BE(16), size);
    assert.equal(png.readUInt32BE(20), size);
  }
});

console.log(JSON.stringify({ ok: true, passed, hosts: ["Word", "Excel", "PowerPoint"], commands: 9, protocolVersion: companion.protocolVersion }));
