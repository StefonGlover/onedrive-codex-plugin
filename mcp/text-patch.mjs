const UTF8_BOM = Buffer.from([0xef, 0xbb, 0xbf]);
const UTF16LE_BOM = Buffer.from([0xff, 0xfe]);
const UTF16BE_BOM = Buffer.from([0xfe, 0xff]);

export function decodeTextBuffer(bytes) {
  let encoding = "utf8";
  let bom = Buffer.alloc(0);
  let body = Buffer.from(bytes);
  if (body.subarray(0, 3).equals(UTF8_BOM)) {
    bom = UTF8_BOM; body = body.subarray(3);
  } else if (body.subarray(0, 2).equals(UTF16LE_BOM)) {
    encoding = "utf16le"; bom = UTF16LE_BOM; body = body.subarray(2);
  } else if (body.subarray(0, 2).equals(UTF16BE_BOM)) {
    encoding = "utf16be"; bom = UTF16BE_BOM; body = body.subarray(2);
  } else if (body.includes(0)) {
    throw new Error("Text patching refused content containing NUL bytes without a supported UTF-16 BOM.");
  }
  if (encoding === "utf16be") {
    const swapped = Buffer.alloc(body.length);
    for (let index = 0; index + 1 < body.length; index += 2) { swapped[index] = body[index + 1]; swapped[index + 1] = body[index]; }
    body = swapped;
  }
  const text = body.toString(encoding === "utf16be" ? "utf16le" : encoding);
  const newline = text.includes("\r\n") ? "\r\n" : text.includes("\r") ? "\r" : "\n";
  return { text, encoding, bom, newline, trailingNewline: /(?:\r\n|\r|\n)$/.test(text) };
}

export function encodeTextBuffer(text, metadata) {
  let body = Buffer.from(text, metadata.encoding === "utf16be" ? "utf16le" : metadata.encoding);
  if (metadata.encoding === "utf16be") {
    const swapped = Buffer.alloc(body.length);
    for (let index = 0; index + 1 < body.length; index += 2) { swapped[index] = body[index + 1]; swapped[index + 1] = body[index]; }
    body = swapped;
  }
  return Buffer.concat([metadata.bom, body]);
}

function splitLines(text) {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}

function applyUnifiedDiff(text, diff) {
  const source = splitLines(text);
  const patchLines = splitLines(diff);
  const output = [];
  let sourceIndex = 0;
  let patchIndex = 0;
  while (patchIndex < patchLines.length && !patchLines[patchIndex].startsWith("@@")) patchIndex += 1;
  if (patchIndex >= patchLines.length) throw new Error("Unified diff contains no hunk header.");
  while (patchIndex < patchLines.length) {
    const header = patchLines[patchIndex++];
    const match = header.match(/^@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@/);
    if (!match) throw new Error(`Invalid unified diff hunk header: ${header}`);
    const oldStart = Number(match[1]) - 1;
    if (oldStart < sourceIndex) throw new Error("Unified diff hunks overlap or are out of order.");
    output.push(...source.slice(sourceIndex, oldStart));
    sourceIndex = oldStart;
    let oldConsumed = 0;
    while (patchIndex < patchLines.length && !patchLines[patchIndex].startsWith("@@")) {
      const line = patchLines[patchIndex++];
      if (line === "\\ No newline at end of file") continue;
      const prefix = line[0];
      const value = line.slice(1);
      if (prefix === " ") {
        if (source[sourceIndex] !== value) throw new Error(`Unified diff context mismatch at source line ${sourceIndex + 1}.`);
        output.push(source[sourceIndex++]); oldConsumed += 1;
      } else if (prefix === "-") {
        if (source[sourceIndex] !== value) throw new Error(`Unified diff deletion mismatch at source line ${sourceIndex + 1}.`);
        sourceIndex += 1; oldConsumed += 1;
      } else if (prefix === "+") output.push(value);
      else if (line !== "") throw new Error(`Invalid unified diff line: ${line}`);
    }
    const expectedOldCount = match[2] === undefined ? 1 : Number(match[2]);
    if (oldConsumed !== expectedOldCount) throw new Error(`Unified diff hunk consumed ${oldConsumed} old lines; expected ${expectedOldCount}.`);
  }
  output.push(...source.slice(sourceIndex));
  return output.join("\n");
}

function pointerTokens(path) {
  if (path === "") return [];
  if (!path.startsWith("/")) throw new Error(`JSON/YAML patch path must be a JSON Pointer: ${path}`);
  return path.slice(1).split("/").map((token) => token.replace(/~1/g, "/").replace(/~0/g, "~"));
}

function parentFor(root, path, create = false) {
  const tokens = pointerTokens(path);
  if (!tokens.length) return { parent: null, key: null, tokens };
  let current = root;
  for (const token of tokens.slice(0, -1)) {
    if (current === null || typeof current !== "object") throw new Error(`Patch path traverses a scalar at ${path}.`);
    if (!Object.hasOwn(current, token)) {
      if (!create) throw new Error(`Patch path does not exist: ${path}`);
      current[token] = {};
    }
    current = current[token];
  }
  return { parent: current, key: tokens.at(-1), tokens };
}

function valueAt(root, path) {
  let current = root;
  for (const token of pointerTokens(path)) {
    if (current === null || typeof current !== "object" || !Object.hasOwn(current, token)) throw new Error(`Patch path does not exist: ${path}`);
    current = current[token];
  }
  return current;
}

function applyJsonPatchValue(root, operations) {
  let document = structuredClone(root);
  for (const operation of operations) {
    if (!operation || typeof operation !== "object" || Array.isArray(operation)) throw new Error("Each JSON/YAML patch operation must be an object.");
    if (!["add", "remove", "replace", "test", "copy", "move"].includes(operation.op)) throw new Error(`Unsupported JSON/YAML patch operation: ${operation.op}`);
    if (operation.op === "test") {
      if (JSON.stringify(valueAt(document, operation.path)) !== JSON.stringify(operation.value)) throw new Error(`Patch test failed at ${operation.path}.`);
      continue;
    }
    let value = operation.value;
    if (operation.op === "copy" || operation.op === "move") value = structuredClone(valueAt(document, operation.from));
    if (operation.op === "move") {
      const source = parentFor(document, operation.from);
      if (source.parent === null) document = null;
      else if (Array.isArray(source.parent)) source.parent.splice(Number(source.key), 1);
      else delete source.parent[source.key];
    }
    const target = parentFor(document, operation.path, operation.op === "add");
    if (target.parent === null) {
      if (operation.op === "remove") throw new Error("Removing the document root is not supported.");
      document = structuredClone(value);
    } else if (Array.isArray(target.parent)) {
      const index = target.key === "-" ? target.parent.length : Number(target.key);
      if (!Number.isInteger(index) || index < 0 || index > target.parent.length) throw new Error(`Invalid array index at ${operation.path}.`);
      if (operation.op === "add" || operation.op === "copy" || operation.op === "move") target.parent.splice(index, 0, structuredClone(value));
      else if (operation.op === "remove") target.parent.splice(index, 1);
      else {
        if (index >= target.parent.length) throw new Error(`Patch path does not exist: ${operation.path}`);
        target.parent[index] = structuredClone(value);
      }
    } else {
      if (operation.op !== "add" && !Object.hasOwn(target.parent, target.key)) throw new Error(`Patch path does not exist: ${operation.path}`);
      if (operation.op === "remove") delete target.parent[target.key];
      else target.parent[target.key] = structuredClone(value);
    }
  }
  return document;
}

function parseScalar(value) {
  const trimmed = value.trim();
  if (trimmed === "" || trimmed === "null" || trimmed === "~") return null;
  if (trimmed === "true") return true;
  if (trimmed === "false") return false;
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(trimmed)) return Number(trimmed);
  if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || (trimmed.startsWith("'") && trimmed.endsWith("'"))) return trimmed.slice(1, -1).replace(trimmed[0] === '"' ? /\\"/g : /''/g, trimmed[0]);
  return trimmed;
}

function parseSafeYaml(text) {
  if (/(^|\s)[&*!]|(^|\s)<<\s*:/.test(text)) throw new Error("YAML anchors, aliases, custom tags, and merge keys are not supported.");
  const lines = splitLines(text).map((raw, index) => ({ raw, index: index + 1 })).filter(({ raw }) => raw.trim() && !raw.trimStart().startsWith("#"));
  const root = {};
  const stack = [{ indent: -1, value: root }];
  for (const { raw, index } of lines) {
    if (/\t/.test(raw.slice(0, raw.length - raw.trimStart().length))) throw new Error(`YAML indentation must use spaces at line ${index}.`);
    const indent = raw.length - raw.trimStart().length;
    const content = raw.trim();
    while (stack.length > 1 && indent <= stack.at(-1).indent) stack.pop();
    const container = stack.at(-1).value;
    const match = content.match(/^([^:#][^:]*?):(?:\s+(.*))?$/);
    if (!match || Array.isArray(container)) throw new Error(`Only safe YAML mappings are supported; invalid line ${index}.`);
    const key = parseScalar(match[1]);
    if (typeof key !== "string" || !key) throw new Error(`Invalid YAML mapping key at line ${index}.`);
    if (Object.hasOwn(container, key)) throw new Error(`Duplicate YAML key '${key}' at line ${index}.`);
    if (match[2] === undefined) {
      container[key] = {};
      stack.push({ indent, value: container[key] });
    } else container[key] = parseScalar(match[2].replace(/\s+#.*$/, ""));
  }
  return root;
}

function yamlScalar(value) {
  if (value === null) return "null";
  if (typeof value === "boolean" || typeof value === "number") return String(value);
  const text = String(value);
  return /^[A-Za-z0-9_.\/-]+$/.test(text) && !["true", "false", "null", "~"].includes(text) ? text : JSON.stringify(text);
}

function stringifySafeYaml(value, indent = 0) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Safe YAML patch output must remain a mapping.");
  const lines = [];
  for (const [key, child] of Object.entries(value)) {
    const prefix = `${" ".repeat(indent)}${key}:`;
    if (child && typeof child === "object" && !Array.isArray(child)) lines.push(prefix, stringifySafeYaml(child, indent + 2));
    else lines.push(`${prefix} ${yamlScalar(child)}`);
  }
  return lines.filter(Boolean).join("\n");
}

function parseCsv(text, delimiter = ",") {
  const rows = [];
  let row = [], field = "", quoted = false;
  const normalizedText = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  for (let index = 0; index < normalizedText.length; index += 1) {
    const char = normalizedText[index];
    if (quoted) {
      if (char === '"' && normalizedText[index + 1] === '"') { field += '"'; index += 1; }
      else if (char === '"') quoted = false;
      else field += char;
    } else if (char === '"' && field === "") quoted = true;
    else if (char === delimiter) { row.push(field); field = ""; }
    else if (char === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
    else field += char;
  }
  if (quoted) throw new Error("CSV contains an unterminated quoted field.");
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function csvCell(value, delimiter) {
  const text = value == null ? "" : String(value);
  return /["\r\n]/.test(text) || text.includes(delimiter) ? `"${text.replace(/"/g, '""')}"` : text;
}

function applyCsvPatch(text, patch) {
  const delimiter = patch.delimiter || ",";
  if (typeof delimiter !== "string" || delimiter.length !== 1) throw new Error("CSV delimiter must be one character.");
  const rows = parseCsv(text, delimiter);
  if (!rows.length) throw new Error("CSV patch requires a header row.");
  const headers = rows[0];
  const keyIndex = headers.indexOf(patch.keyColumn);
  if (keyIndex < 0) throw new Error(`CSV key column not found: ${patch.keyColumn}`);
  const records = rows.slice(1);
  for (const operation of patch.operations || []) {
    const values = Array.isArray(operation.values)
      ? Object.fromEntries(operation.values.map((entry) => [entry.column, entry.value]))
      : operation.values || {};
    const matches = records.map((row, index) => ({ row, index })).filter(({ row }) => row[keyIndex] === String(operation.key));
    if (operation.op === "insert") {
      if (matches.length) throw new Error(`CSV key already exists: ${operation.key}`);
      const row = headers.map((header) => values[header] == null ? "" : String(values[header]));
      row[keyIndex] = String(operation.key); records.push(row);
    } else {
      if (matches.length !== 1) throw new Error(`CSV key must resolve exactly once: ${operation.key}`);
      if (operation.op === "delete") records.splice(matches[0].index, 1);
      else if (operation.op === "update") for (const [header, value] of Object.entries(values)) {
        const column = headers.indexOf(header);
        if (column < 0) throw new Error(`CSV column not found: ${header}`);
        matches[0].row[column] = value == null ? "" : String(value);
      } else throw new Error(`Unsupported CSV patch operation: ${operation.op}`);
    }
  }
  return [headers, ...records].map((row) => row.map((cell) => csvCell(cell, delimiter)).join(delimiter)).join("\n");
}

export function applyTextPatch(bytes, patch) {
  const metadata = decodeTextBuffer(bytes);
  let output;
  if (patch.mode === "unified") output = applyUnifiedDiff(metadata.text, patch.diff);
  else if (patch.mode === "json") output = `${JSON.stringify(applyJsonPatchValue(JSON.parse(metadata.text), patch.operations), null, patch.indent ?? 2)}`;
  else if (patch.mode === "yaml") output = stringifySafeYaml(applyJsonPatchValue(parseSafeYaml(metadata.text), patch.operations));
  else if (patch.mode === "csv") output = applyCsvPatch(metadata.text, patch);
  else throw new Error(`Unsupported text patch mode: ${patch.mode}`);
  output = output.replace(/\n/g, metadata.newline);
  if (metadata.trailingNewline && !output.endsWith(metadata.newline)) output += metadata.newline;
  if (!metadata.trailingNewline && output.endsWith(metadata.newline)) output = output.slice(0, -metadata.newline.length);
  return {
    bytes: encodeTextBuffer(output, metadata),
    beforeText: metadata.text,
    afterText: output,
    encoding: metadata.encoding,
    bom: metadata.bom.length > 0,
    newline: metadata.newline === "\r\n" ? "CRLF" : metadata.newline === "\r" ? "CR" : "LF",
    trailingNewline: metadata.trailingNewline
  };
}

export function boundedLineDiff(before, after, maxChanges = 200) {
  const left = splitLines(before), right = splitLines(after);
  const changes = [];
  const count = Math.max(left.length, right.length);
  for (let index = 0; index < count && changes.length < maxChanges; index += 1) {
    if (left[index] !== right[index]) changes.push({ line: index + 1, before: left[index] ?? null, after: right[index] ?? null });
  }
  return { changes, changeCount: changes.length, truncated: changes.length === maxChanges && count > maxChanges, beforeLines: left.length, afterLines: right.length };
}
