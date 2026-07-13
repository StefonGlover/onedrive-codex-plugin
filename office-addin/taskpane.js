/* global Office, Word, Excel, PowerPoint */
(() => {
  "use strict";

  const protocolVersion = "codex-office-companion/1";
  const companionVersion = "1.1.1";
  const maxCommandBytes = 64 * 1024;
  const hostDefinitions = {
    Word: { requirementSet: { name: "WordApi", version: "1.3" }, commands: ["replaceText", "setSelectedText", "insertParagraph"] },
    Excel: { requirementSet: { name: "ExcelApi", version: "1.7" }, commands: ["setRange", "clearRange", "formatRange"] },
    PowerPoint: { requirementSet: { name: "PowerPointApi", version: "1.5" }, commands: ["setSelectedText", "setSelectedTextStyle", "deleteSelectedShapes"] }
  };
  const starterCommands = {
    Word: { type: "replaceText", find: "old text", replace: "new text" },
    Excel: { type: "setRange", sheet: "Sheet1", address: "A1:B2", values: [[1, 2], [3, 4]] },
    PowerPoint: { type: "setSelectedTextStyle", bold: true }
  };
  const excelClearModes = ["All", "Contents", "Formats", "Hyperlinks", "RemoveHyperlinks"];
  const byId = (id) => document.getElementById(id);
  const has = (value, key) => Object.hasOwn(value, key);
  const result = (value) => { byId("result").textContent = typeof value === "string" ? value : JSON.stringify(value, null, 2); };
  const activeHost = () => Office.context.host;
  const normalizedHost = (host = activeHost()) => Object.entries(Office.HostType).find(([, value]) => value === host)?.[0] || String(host || "Unknown");
  const requirementSupported = (definition) => Office.context.requirements.isSetSupported(definition.requirementSet.name, definition.requirementSet.version);

  function assertKeys(command, allowed) {
    for (const key of Object.keys(command)) if (!allowed.includes(key)) throw new Error(`Unknown command property: ${key}`);
  }

  function assertString(command, key, { required = false, nonEmpty = false } = {}) {
    if (!has(command, key)) {
      if (required) throw new Error(`${key} is required.`);
      return;
    }
    if (typeof command[key] !== "string") throw new Error(`${key} must be a string.`);
    if (nonEmpty && command[key].trim().length === 0) throw new Error(`${key} must not be empty.`);
  }

  function assertBoolean(command, key) {
    if (has(command, key) && typeof command[key] !== "boolean") throw new Error(`${key} must be a boolean.`);
  }

  function assertFiniteNumber(command, key, { min = -Infinity, max = Infinity } = {}) {
    if (!has(command, key)) return;
    const value = command[key];
    if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(`${key} must be a finite number.`);
    if (value < min || value > max) throw new Error(`${key} must be between ${min} and ${max}.`);
  }

  function assertEnum(command, key, values) {
    if (has(command, key) && !values.includes(command[key])) throw new Error(`${key} must be one of: ${values.join(", ")}.`);
  }

  function assertMatrix(command, key, cellValidator, cellDescription) {
    if (!has(command, key)) return null;
    const matrix = command[key];
    if (!Array.isArray(matrix) || matrix.length === 0) throw new Error(`${key} must be a non-empty two-dimensional array.`);
    let columns = null;
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex];
      if (!Array.isArray(row) || row.length === 0) throw new Error(`${key}[${rowIndex}] must be a non-empty array.`);
      if (columns === null) columns = row.length;
      if (row.length !== columns) throw new Error(`${key} must be rectangular.`);
      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        if (!cellValidator(row[columnIndex])) throw new Error(`${key}[${rowIndex}][${columnIndex}] must be ${cellDescription}.`);
      }
    }
    return { rows: matrix.length, columns };
  }

  function assertMatrixDimensions(key, dimensions, range) {
    if (!dimensions) return;
    if (dimensions.rows !== range.rowCount || dimensions.columns !== range.columnCount) {
      throw new Error(`${key} dimensions ${dimensions.rows}x${dimensions.columns} must match range dimensions ${range.rowCount}x${range.columnCount}.`);
    }
  }

  function assertSerializedCommandSize(value) {
    let serialized;
    try {
      serialized = JSON.stringify(value);
    } catch {
      throw new Error("Command must be JSON-serializable.");
    }
    if (serialized === undefined) throw new Error("Command must be JSON-serializable.");
    if (new TextEncoder().encode(serialized).length > maxCommandBytes) throw new Error(`Command exceeds ${maxCommandBytes} bytes.`);
  }

  function getCapabilities(info = {}) {
    const host = normalizedHost(info.host);
    const definition = hostDefinitions[host];
    const supported = Boolean(definition && requirementSupported(definition));
    return {
      protocolVersion,
      companionVersion,
      host,
      platform: info.platform || Office.context.platform || "Unknown",
      available: supported,
      requirementSet: definition ? { ...definition.requirementSet, supported } : null,
      commands: supported ? [...definition.commands] : [],
      limits: { maxCommandBytes, oneActiveDocument: true },
      transport: { manualPaste: true, remoteCommands: false, telemetry: false }
    };
  }

  function normalizeEnvelope(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Command must be a JSON object.");
    const command = has(value, "command") ? (() => {
      assertKeys(value, ["protocolVersion", "command"]);
      if (value.protocolVersion !== protocolVersion) throw new Error(`Unsupported companion protocolVersion: ${value.protocolVersion}`);
      if (!value.command || typeof value.command !== "object" || Array.isArray(value.command)) throw new Error("command must be a JSON object.");
      return value.command;
    })() : value;
    assertString(command, "type", { required: true, nonEmpty: true });
    return command;
  }

  function validateWordCommand(command) {
    if (!hostDefinitions.Word.commands.includes(command.type)) throw new Error("Unsupported Word companion command.");
    if (command.type === "replaceText") {
      assertKeys(command, ["type", "find", "replace", "all", "matchCase"]);
      assertString(command, "find", { required: true, nonEmpty: true });
      assertString(command, "replace");
      assertBoolean(command, "all");
      assertBoolean(command, "matchCase");
      return;
    }
    if (command.type === "setSelectedText") {
      assertKeys(command, ["type", "text"]);
      assertString(command, "text", { required: true });
      return;
    }
    assertKeys(command, ["type", "text", "location"]);
    assertString(command, "text", { required: true });
    assertEnum(command, "location", ["before", "after"]);
  }

  async function applyWord(command) {
    validateWordCommand(command);
    if (!requirementSupported(hostDefinitions.Word)) throw new Error("WordApi 1.3 is required.");
    if (command.type === "replaceText") {
      return Word.run(async (context) => {
        const matches = context.document.body.search(command.find, { matchCase: command.matchCase === true, matchWholeWord: false });
        matches.load("items");
        await context.sync();
        const selected = command.all === false ? matches.items.slice(0, 1) : matches.items;
        for (const range of selected) range.insertText(has(command, "replace") ? command.replace : "", Word.InsertLocation.replace);
        await context.sync();
        return { host: "Word", operation: command.type, changed: selected.length };
      });
    }
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (command.type === "setSelectedText") {
        selection.load("isEmpty");
        await context.sync();
        if (selection.isEmpty) throw new Error("setSelectedText requires a non-empty Word text selection.");
        selection.insertText(command.text, Word.InsertLocation.replace);
      } else {
        const location = command.location === "before" ? Word.InsertLocation.before : Word.InsertLocation.after;
        selection.insertParagraph(command.text, location);
      }
      await context.sync();
      return { host: "Word", operation: command.type, changed: 1 };
    });
  }

  function validateExcelCommand(command) {
    if (!hostDefinitions.Excel.commands.includes(command.type)) throw new Error("Unsupported Excel companion command.");
    const allowed = command.type === "setRange" ? ["type", "sheet", "address", "values", "formulas", "numberFormat"]
      : command.type === "clearRange" ? ["type", "sheet", "address", "applyTo"]
        : ["type", "sheet", "address", "fillColor", "fontColor", "fontName", "fontSize", "bold", "italic", "numberFormat"];
    assertKeys(command, allowed);
    assertString(command, "sheet", { required: true, nonEmpty: true });
    assertString(command, "address", { required: true, nonEmpty: true });

    if (command.type === "setRange") {
      if (!has(command, "values") && !has(command, "formulas") && !has(command, "numberFormat")) throw new Error("setRange requires values, formulas, or numberFormat.");
    } else if (command.type === "clearRange") {
      assertEnum(command, "applyTo", excelClearModes);
    } else {
      const styleKeys = allowed.slice(3);
      if (!styleKeys.some((key) => has(command, key))) throw new Error("formatRange requires at least one formatting property.");
      for (const key of ["fillColor", "fontColor", "fontName"]) assertString(command, key, { nonEmpty: true });
      assertFiniteNumber(command, "fontSize", { min: 1, max: 409 });
      assertBoolean(command, "bold");
      assertBoolean(command, "italic");
    }

    const cellValue = (value) => value === null || typeof value === "string" || typeof value === "boolean" || (typeof value === "number" && Number.isFinite(value));
    return {
      values: assertMatrix(command, "values", cellValue, "a string, boolean, finite number, or null"),
      formulas: assertMatrix(command, "formulas", cellValue, "a string, boolean, finite number, or null"),
      numberFormat: assertMatrix(command, "numberFormat", (value) => typeof value === "string", "a string")
    };
  }

  async function applyExcel(command) {
    const matrices = validateExcelCommand(command);
    if (!requirementSupported(hostDefinitions.Excel)) throw new Error("ExcelApi 1.7 is required.");
    return Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(command.sheet).getRange(command.address);
      range.load("rowCount,columnCount");
      await context.sync();
      for (const [key, dimensions] of Object.entries(matrices)) assertMatrixDimensions(key, dimensions, range);

      if (command.type === "setRange") {
        if (has(command, "values")) range.values = command.values;
        if (has(command, "formulas")) range.formulas = command.formulas;
        if (has(command, "numberFormat")) range.numberFormat = command.numberFormat;
      } else if (command.type === "clearRange") {
        range.clear(has(command, "applyTo") ? command.applyTo : "Contents");
      } else {
        if (has(command, "fillColor")) range.format.fill.color = command.fillColor;
        if (has(command, "fontColor")) range.format.font.color = command.fontColor;
        if (has(command, "fontName")) range.format.font.name = command.fontName;
        if (has(command, "fontSize")) range.format.font.size = command.fontSize;
        if (has(command, "bold")) range.format.font.bold = command.bold;
        if (has(command, "italic")) range.format.font.italic = command.italic;
        if (has(command, "numberFormat")) range.numberFormat = command.numberFormat;
      }
      await context.sync();
      return { host: "Excel", operation: command.type, sheet: command.sheet, address: command.address, changed: 1 };
    });
  }

  function validatePowerPointCommand(command) {
    if (!hostDefinitions.PowerPoint.commands.includes(command.type)) throw new Error("Unsupported PowerPoint companion command.");
    const allowed = command.type === "setSelectedText" ? ["type", "text"]
      : command.type === "setSelectedTextStyle" ? ["type", "fontName", "fontSize", "color", "bold", "italic", "underline"]
        : ["type"];
    assertKeys(command, allowed);
    if (command.type === "setSelectedText") {
      assertString(command, "text", { required: true });
    } else if (command.type === "setSelectedTextStyle") {
      const styleKeys = allowed.slice(1);
      if (!styleKeys.some((key) => has(command, key))) throw new Error("setSelectedTextStyle requires at least one formatting property.");
      for (const key of ["fontName", "color"]) assertString(command, key, { nonEmpty: true });
      assertFiniteNumber(command, "fontSize", { min: 1, max: 400 });
      for (const key of ["bold", "italic", "underline"]) assertBoolean(command, key);
    }
  }

  async function applyPowerPoint(command) {
    validatePowerPointCommand(command);
    if (!requirementSupported(hostDefinitions.PowerPoint)) throw new Error("PowerPointApi 1.5 is required.");
    return PowerPoint.run(async (context) => {
      if (command.type === "deleteSelectedShapes") {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items");
        await context.sync();
        if (shapes.items.length === 0) throw new Error("deleteSelectedShapes requires at least one selected shape.");
        for (const shape of shapes.items) shape.delete();
        await context.sync();
        return { host: "PowerPoint", operation: command.type, changed: shapes.items.length };
      }

      const textRange = context.presentation.getSelectedTextRangeOrNullObject();
      textRange.load("isNullObject");
      await context.sync();
      if (textRange.isNullObject) throw new Error(`${command.type} requires a non-empty PowerPoint text selection.`);
      if (command.type === "setSelectedText") {
        textRange.text = command.text;
      } else {
        if (has(command, "fontName")) textRange.font.name = command.fontName;
        if (has(command, "fontSize")) textRange.font.size = command.fontSize;
        if (has(command, "color")) textRange.font.color = command.color;
        if (has(command, "bold")) textRange.font.bold = command.bold;
        if (has(command, "italic")) textRange.font.italic = command.italic;
        if (has(command, "underline")) {
          textRange.font.underline = command.underline
            ? PowerPoint.ShapeFontUnderlineStyle.single
            : PowerPoint.ShapeFontUnderlineStyle.none;
        }
      }
      await context.sync();
      return { host: "PowerPoint", operation: command.type, changed: 1 };
    });
  }

  async function executeCommand(value) {
    assertSerializedCommandSize(value);
    const command = normalizeEnvelope(value);
    const host = normalizedHost();
    const definition = hostDefinitions[host];
    if (!definition || !requirementSupported(definition)) throw new Error("The active Office host does not satisfy the companion requirement set.");
    return host === "Word" ? applyWord(command) : host === "Excel" ? applyExcel(command) : applyPowerPoint(command);
  }

  async function apply() {
    byId("apply").disabled = true;
    try {
      result({ ok: true, ...(await executeCommand(JSON.parse(byId("command").value))) });
    } catch (error) {
      result({ ok: false, error: String(error?.message || error) });
    } finally {
      byId("apply").disabled = false;
    }
  }

  globalThis.CodexOfficeCompanion = Object.freeze({ protocolVersion, getCapabilities, executeCommand });
  Office.onReady((info) => {
    const capabilities = getCapabilities(info);
    byId("host").textContent = capabilities.available
      ? `${capabilities.host} on ${capabilities.platform}; ${capabilities.requirementSet.name} ${capabilities.requirementSet.version}; ${capabilities.commands.length} typed commands`
      : `${capabilities.host} on ${capabilities.platform}; required Office API set unavailable`;
    byId("apply").disabled = !capabilities.available;
    byId("apply").addEventListener("click", apply);
    if (starterCommands[capabilities.host]) byId("command").value = JSON.stringify(starterCommands[capabilities.host], null, 2);
    result(capabilities);
  });
})();
