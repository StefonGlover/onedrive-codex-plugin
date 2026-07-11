/* global Office, Word, Excel, PowerPoint */
(() => {
  "use strict";

  const protocolVersion = "codex-office-companion/1";
  const maxCommandBytes = 64 * 1024;
  const hostDefinitions = {
    Word: { requirementSet: { name: "WordApi", version: "1.3" }, commands: ["replaceText", "setSelectedText", "insertParagraph"] },
    Excel: { requirementSet: { name: "ExcelApi", version: "1.7" }, commands: ["setRange", "clearRange", "formatRange"] },
    PowerPoint: { requirementSet: { name: "PowerPointApi", version: "1.5" }, commands: ["setSelectedText", "setSelectedTextStyle", "deleteSelectedShapes"] }
  };
  const byId = (id) => document.getElementById(id);
  const result = (value) => { byId("result").textContent = typeof value === "string" ? value : JSON.stringify(value, null, 2); };
  const assertKeys = (command, allowed) => {
    for (const key of Object.keys(command)) if (!allowed.includes(key)) throw new Error(`Unknown command property: ${key}`);
  };
  const activeHost = () => Office.context.host;
  const normalizedHost = (host = activeHost()) => Object.entries(Office.HostType).find(([, value]) => value === host)?.[0] || String(host || "Unknown");
  const requirementSupported = (definition) => Office.context.requirements.isSetSupported(definition.requirementSet.name, definition.requirementSet.version);

  function getCapabilities(info = {}) {
    const host = normalizedHost(info.host);
    const definition = hostDefinitions[host];
    const supported = Boolean(definition && requirementSupported(definition));
    return {
      protocolVersion,
      companionVersion: "1.1.0",
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
    if (!Object.hasOwn(value, "command")) return value;
    assertKeys(value, ["protocolVersion", "command"]);
    if (value.protocolVersion !== protocolVersion) throw new Error(`Unsupported companion protocolVersion: ${value.protocolVersion}`);
    if (!value.command || typeof value.command !== "object" || Array.isArray(value.command)) throw new Error("command must be a JSON object.");
    return value.command;
  }

  async function applyWord(command) {
    if (!hostDefinitions.Word.commands.includes(command.type)) throw new Error("Unsupported Word companion command.");
    if (!requirementSupported(hostDefinitions.Word)) throw new Error("WordApi 1.3 is required.");
    if (command.type === "replaceText") {
      assertKeys(command, ["type", "find", "replace", "all", "matchCase"]);
      if (!command.find) throw new Error("find is required.");
      return Word.run(async (context) => {
        const matches = context.document.body.search(command.find, { matchCase: command.matchCase === true, matchWholeWord: false });
        matches.load("items");
        await context.sync();
        const selected = command.all === false ? matches.items.slice(0, 1) : matches.items;
        for (const range of selected) range.insertText(String(command.replace ?? ""), Word.InsertLocation.replace);
        await context.sync();
        return { host: "Word", operation: command.type, changed: selected.length };
      });
    }
    assertKeys(command, command.type === "setSelectedText" ? ["type", "text"] : ["type", "text", "location"]);
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (command.type === "setSelectedText") selection.insertText(String(command.text ?? ""), Word.InsertLocation.replace);
      else selection.insertParagraph(String(command.text ?? ""), command.location === "before" ? Word.InsertLocation.before : Word.InsertLocation.after);
      await context.sync();
      return { host: "Word", operation: command.type, changed: 1 };
    });
  }

  async function applyExcel(command) {
    if (!hostDefinitions.Excel.commands.includes(command.type)) throw new Error("Unsupported Excel companion command.");
    if (!requirementSupported(hostDefinitions.Excel)) throw new Error("ExcelApi 1.7 is required.");
    const allowed = command.type === "setRange" ? ["type", "sheet", "address", "values", "formulas", "numberFormat"]
      : command.type === "clearRange" ? ["type", "sheet", "address", "applyTo"]
        : ["type", "sheet", "address", "fillColor", "fontColor", "fontName", "fontSize", "bold", "italic", "numberFormat"];
    assertKeys(command, allowed);
    if (!command.sheet || !command.address) throw new Error("sheet and address are required.");
    return Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(command.sheet).getRange(command.address);
      if (command.type === "setRange") {
        if (command.values === undefined && command.formulas === undefined && command.numberFormat === undefined) throw new Error("setRange requires values, formulas, or numberFormat.");
        if (command.values !== undefined) range.values = command.values;
        if (command.formulas !== undefined) range.formulas = command.formulas;
        if (command.numberFormat !== undefined) range.numberFormat = command.numberFormat;
      } else if (command.type === "clearRange") {
        range.clear(command.applyTo || "Contents");
      } else {
        if (command.fillColor !== undefined) range.format.fill.color = command.fillColor;
        if (command.fontColor !== undefined) range.format.font.color = command.fontColor;
        if (command.fontName !== undefined) range.format.font.name = command.fontName;
        if (command.fontSize !== undefined) range.format.font.size = Number(command.fontSize);
        if (command.bold !== undefined) range.format.font.bold = command.bold === true;
        if (command.italic !== undefined) range.format.font.italic = command.italic === true;
        if (command.numberFormat !== undefined) range.numberFormat = command.numberFormat;
      }
      await context.sync();
      return { host: "Excel", operation: command.type, sheet: command.sheet, address: command.address, changed: 1 };
    });
  }

  async function applyPowerPoint(command) {
    if (!hostDefinitions.PowerPoint.commands.includes(command.type)) throw new Error("Unsupported PowerPoint companion command.");
    if (!requirementSupported(hostDefinitions.PowerPoint)) throw new Error("PowerPointApi 1.5 is required.");
    const allowed = command.type === "setSelectedText" ? ["type", "text"]
      : command.type === "setSelectedTextStyle" ? ["type", "fontName", "fontSize", "color", "bold", "italic", "underline"]
        : ["type"];
    assertKeys(command, allowed);
    return PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      for (const shape of shapes.items) shape.load("textFrame/hasText,textFrame/textRange/text");
      await context.sync();
      let changed = 0;
      for (const shape of shapes.items) {
        if (command.type === "deleteSelectedShapes") {
          shape.delete();
          changed += 1;
          continue;
        }
        if (!shape.textFrame.hasText) continue;
        const textRange = shape.textFrame.textRange;
        if (command.type === "setSelectedText") textRange.text = String(command.text ?? "");
        else {
          if (command.fontName !== undefined) textRange.font.name = command.fontName;
          if (command.fontSize !== undefined) textRange.font.size = Number(command.fontSize);
          if (command.color !== undefined) textRange.font.color = command.color;
          if (command.bold !== undefined) textRange.font.bold = command.bold === true;
          if (command.italic !== undefined) textRange.font.italic = command.italic === true;
          if (command.underline !== undefined) textRange.font.underline = command.underline === true;
        }
        changed += 1;
      }
      await context.sync();
      return { host: "PowerPoint", operation: command.type, changed };
    });
  }

  async function executeCommand(value) {
    const command = normalizeEnvelope(value);
    const definition = hostDefinitions[normalizedHost()];
    if (!definition || !requirementSupported(definition)) throw new Error("The active Office host does not satisfy the companion requirement set.");
    return normalizedHost() === "Word" ? applyWord(command) : normalizedHost() === "Excel" ? applyExcel(command) : applyPowerPoint(command);
  }

  async function apply() {
    byId("apply").disabled = true;
    try {
      const raw = byId("command").value;
      if (new TextEncoder().encode(raw).length > maxCommandBytes) throw new Error(`Command exceeds ${maxCommandBytes} bytes.`);
      result({ ok: true, ...(await executeCommand(JSON.parse(raw))) });
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
    result(capabilities);
  });
})();
