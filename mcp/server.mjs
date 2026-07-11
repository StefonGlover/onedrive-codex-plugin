#!/usr/bin/env node

import { execFileSync, spawn } from "node:child_process";
import { createReadStream, createWriteStream, readFileSync } from "node:fs";
import { appendFile, chmod, copyFile, lstat, mkdir, open, readFile, readdir, realpath, rename as renameFile, rm, stat, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join, parse, relative, resolve } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { createHash, randomUUID } from "node:crypto";
import { AsyncLocalStorage } from "node:async_hooks";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const pluginManifest = JSON.parse(readFileSync(join(pluginRoot, ".codex-plugin", "plugin.json"), "utf8"));
const defaultStorageRoot = join(homedir(), ".codex", "onedrive-plugin");
const configPath = join(defaultStorageRoot, "config.json");
const localConfig = readLocalConfig();
const storageRoot = resolve(process.env.ONEDRIVE_STORAGE_ROOT || localConfig.storageRoot || defaultStorageRoot);
const downloadRoot = join(storageRoot, "downloads");
const cacheRoot = resolve(process.env.ONEDRIVE_CACHE_ROOT || localConfig.cacheRoot || join(storageRoot, "cache"));
const cachePath = join(cacheRoot, "metadata-cache.json");
const contentIndexPath = join(cacheRoot, "content-index.json");
const metadataCacheLockPath = join(cacheRoot, "metadata-cache.lock");
const contentIndexLockPath = join(cacheRoot, "content-index.lock");
const updateRoot = join(storageRoot, "updates");
const backupRoot = join(storageRoot, "backups");
const auditRoot = join(storageRoot, "audit");
const auditPath = join(auditRoot, "mutations.jsonl");
const auditLockPath = join(auditRoot, "mutations.lock");
const officeEditingRoot = join(storageRoot, "office-editing");
const officeHelperPath = join(pluginRoot, "scripts", "office-openxml.py");
const officePythonPath = process.env.ONEDRIVE_OFFICE_PYTHON || localConfig.officeEditing?.pythonPath || "/usr/bin/python3";
const maxOfficePackageBytes = 250 * 1024 * 1024;
const localOneDriveSyncRoots = [
  { path: join(homedir(), "Library", "CloudStorage", "OneDrive"), prefix: false },
  { path: join(homedir(), "Library", "CloudStorage", "OneDrive-"), prefix: true },
  { path: join(homedir(), "OneDrive"), prefix: false },
  { path: join(homedir(), "OneDrive - "), prefix: true }
];
const textFileLimit = 5 * 1024 * 1024;
const maxTextFileReadLimit = 10 * 1024 * 1024;
const defaultMaxIndexedFileSize = 512 * 1024;
const simpleUploadLimit = 250 * 1024 * 1024;
const uploadChunkUnit = 320 * 1024;
const defaultUploadChunkSize = 10 * 1024 * 1024;
const maxUploadChunkSize = 60 * 1024 * 1024;
const defaultSelect = "id,name,size,folder,file,webUrl,createdDateTime,lastModifiedDateTime,parentReference,eTag,cTag,deleted";
const textExtensions = new Set([
  ".bat", ".c", ".cfg", ".conf", ".cpp", ".cs", ".css", ".csv", ".env", ".go", ".h", ".hpp", ".htm",
  ".html", ".ini", ".java", ".js", ".json", ".jsx", ".log", ".md", ".mjs", ".php", ".properties",
  ".py", ".rb", ".rs", ".sh", ".sql", ".svg", ".toml", ".ts", ".tsx", ".txt", ".xml", ".yaml", ".yml"
]);
const textMimePrefixes = ["text/"];
const textMimeTypes = new Set([
  "application/csv",
  "application/ecmascript",
  "application/javascript",
  "application/json",
  "application/sql",
  "application/x-javascript",
  "application/x-ndjson",
  "application/xml",
  "image/svg+xml"
]);
const officeKinds = {
  excel: {
    label: "Excel",
    extensions: new Set([".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods"]),
    mimeTypes: new Set([
      "application/vnd.ms-excel",
      "application/vnd.ms-excel.sheet.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.oasis.opendocument.spreadsheet",
      "text/csv"
    ])
  },
  word: {
    label: "Word",
    extensions: new Set([".docx", ".doc", ".docm", ".rtf", ".odt"]),
    mimeTypes: new Set([
      "application/msword",
      "application/rtf",
      "application/vnd.ms-word.document.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/vnd.oasis.opendocument.text",
      "text/rtf"
    ])
  },
  powerpoint: {
    label: "PowerPoint",
    extensions: new Set([".pptx", ".ppt", ".pptm", ".ppsx", ".odp"]),
    mimeTypes: new Set([
      "application/vnd.ms-powerpoint",
      "application/vnd.ms-powerpoint.presentation.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
      "application/vnd.oasis.opendocument.presentation"
    ])
  }
};

let tokenCache = null;
let pendingDevice = null;
let tokenRefreshPromise = null;
let authGeneration = 0;
let deviceLoginGeneration = 0;
let metadataCacheMemory = null;
let contentIndexMemory = null;
let metadataCacheLoadPromise = null;
let contentIndexLoadPromise = null;
let metadataCacheFileVersion = null;
let contentIndexFileVersion = null;
let metadataMutationQueue = Promise.resolve();
let contentIndexMutationQueue = Promise.resolve();

const previewTokens = new Map();
const previewTokenTtlMs = 15 * 60 * 1000;
const partialBatchMutationWarning = "Batch live mutations are preflighted but not atomic; if a later item fails, earlier remote changes may already have taken effect.";
const toolCallContext = new AsyncLocalStorage();

const outputFormatSchema = {
  type: "string",
  enum: ["compact", "full"],
  default: "compact",
  description: "compact returns chat-friendly item summaries; full returns richer item metadata."
};
const presetSchema = {
  type: "string",
  description: "Optional configured path preset, such as documents, desktop, pictures, or screenshots."
};
const relativePathSchema = {
  type: "string",
  description: "Path appended below the selected preset."
};
const previewTokenSchema = {
  type: "string",
  description: "Token returned by the immediately preceding dry-run preview for this exact high-risk operation."
};
const officeTargetProperties = {
  path: { type: "string", description: "Office file path relative to OneDrive root." },
  itemId: { type: "string", description: "Office file drive item ID." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const officeStructuredSearchProperties = {
  searchText: { type: "string", minLength: 1, description: "Optional in-package text/value query; returns bounded structured matches alongside the normal read result." },
  matchCase: { type: "boolean", default: false },
  maxMatches: { type: "integer", minimum: 1, maximum: 5000, default: 200 }
};
const officeBatchCommonProperties = {
  ...officeTargetProperties,
  operations: {
    type: "array",
    minItems: 1,
    maxItems: 100,
    items: { type: "object", additionalProperties: true },
    description: "Typed Office edit operations executed in order. Use the corresponding read tool first to ground indexes, sheet names, slide indexes, and shape IDs."
  },
  dryRun: { type: "boolean", default: true },
  confirmed: { type: "boolean", default: false },
  expectedName: { type: "string", description: "Safety check for the resolved remote file name." },
  expectedId: { type: "string", description: "Safety check for the resolved remote drive item ID." },
  previewToken: previewTokenSchema,
  createBackup: { type: "boolean", default: true },
  verify: { type: "boolean", default: true },
  allowMacros: { type: "boolean", default: false, description: "Allow editing a macro-enabled package while preserving VBA parts. Never executes macros." },
  allowSignedPackage: { type: "boolean", default: false, description: "Reserved compatibility flag. Digitally signed packages are always refused because native edits invalidate signatures." }
};
const operationObject = (required, properties) => ({ type: "object", required, properties, additionalProperties: false });
const replaceTextOperation = operationObject(["type", "find", "replace"], {
  type: { const: "replaceText" }, find: { type: "string", minLength: 1 }, replace: { type: "string" }, matchCase: { type: "boolean" }, all: { type: "boolean" }
});
const wordOperationsSchema = {
  type: "array", minItems: 1, maxItems: 100, items: { anyOf: [
    { ...replaceTextOperation, properties: { ...replaceTextOperation.properties, scope: { type: "string", enum: ["document", "headers", "footers", "all"] } } },
    operationObject(["type", "paragraphIndex", "text"], { type: { const: "setParagraphText" }, paragraphIndex: { type: "integer", minimum: 0 }, text: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "paragraphIndex", "style"], { type: { const: "setParagraphStyle" }, paragraphIndex: { type: "integer", minimum: 0 }, style: { type: "string", minLength: 1 }, part: { type: "string" } }),
    operationObject(["type", "text"], { type: { const: "insertParagraph" }, afterIndex: { type: "integer", minimum: 0 }, text: { type: "string" }, style: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "tableIndex", "rowIndex", "columnIndex", "text"], { type: { const: "setTableCell" }, tableIndex: { type: "integer", minimum: 0 }, rowIndex: { type: "integer", minimum: 0 }, columnIndex: { type: "integer", minimum: 0 }, text: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "text"], { type: { const: "setContentControlText" }, contentControlIndex: { type: "integer", minimum: 0 }, id: { type: "string" }, tag: { type: "string" }, text: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "paragraphIndex", "text", "url"], { type: { const: "addHyperlink" }, paragraphIndex: { type: "integer", minimum: 0 }, text: { type: "string", minLength: 1 }, url: { type: "string", minLength: 1 }, part: { type: "string" } }),
    operationObject(["type", "paragraphIndex", "text"], { type: { const: "addComment" }, paragraphIndex: { type: "integer", minimum: 0 }, text: { type: "string", minLength: 1 }, author: { type: "string", maxLength: 255 }, initials: { type: "string", maxLength: 16 } }),
    operationObject(["type", "rows"], { type: { const: "insertTable" }, afterParagraphIndex: { type: "integer", minimum: 0 }, rows: { type: "array", minItems: 1, maxItems: 100, items: { type: "array", minItems: 1, maxItems: 50, items: { type: "string" } } }, style: { type: "string", minLength: 1 } })
  ] }
};
const excelBaseOperation = { sheet: { type: "string", minLength: 1 }, address: { type: "string", pattern: "^[A-Za-z]{1,3}[1-9][0-9]*(?::[A-Za-z]{1,3}[1-9][0-9]*)?$" } };
const excelOperationsSchema = {
  type: "array", minItems: 1, maxItems: 100, items: { anyOf: [
    operationObject(["type", "sheet", "address"], { type: { const: "setCell" }, ...excelBaseOperation, value: {} }),
    operationObject(["type", "sheet", "address", "formula"], { type: { const: "setFormula" }, ...excelBaseOperation, formula: { type: "string" }, value: {} }),
    operationObject(["type", "sheet", "address"], { type: { const: "setRange" }, ...excelBaseOperation, values: { type: "array", items: { type: "array" } }, formulas: { type: "array", items: { type: "array" } } }),
    operationObject(["type", "sheet", "address"], { type: { const: "clearRange" }, ...excelBaseOperation, contents: { type: "boolean" }, format: { type: "boolean" } }),
    operationObject(["type", "sheet", "address", "styleIndex"], { type: { const: "setStyle" }, ...excelBaseOperation, styleIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "sheet", "address", "formatCode"], { type: { const: "setNumberFormat" }, ...excelBaseOperation, formatCode: { type: "string", minLength: 1, maxLength: 255 } }),
    operationObject(["type", "sheet", "address", "formula", "fillColor"], { type: { const: "addConditionalFormat" }, ...excelBaseOperation, ruleType: { type: "string", enum: ["expression", "cellIs"] }, operator: { type: "string", enum: ["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"] }, formula: { type: "string", minLength: 1 }, fillColor: { type: "string", pattern: "^[0-9A-Fa-f]{6}$" } }),
    operationObject(["type", "sheet", "address", "validationType"], { type: { const: "setDataValidation" }, ...excelBaseOperation, validationType: { type: "string", enum: ["whole", "decimal", "list", "date", "time", "textLength", "custom"] }, operator: { type: "string" }, formula1: { type: "string" }, formula2: { type: "string" }, allowBlank: { type: "boolean" } }),
    operationObject(["type", "sheet"], { type: { const: "freezePanes" }, sheet: { type: "string", minLength: 1 }, rows: { type: "integer", minimum: 0, maximum: 1048575 }, columns: { type: "integer", minimum: 0, maximum: 16383 } }),
    operationObject(["type", "sheet", "address", "width"], { type: { const: "setColumnWidth" }, ...excelBaseOperation, width: { type: "number", exclusiveMinimum: 0, maximum: 255 } }),
    operationObject(["type", "table", "values"], { type: { const: "addTableRow" }, table: { type: "string", minLength: 1 }, index: { type: ["integer", "null"], minimum: 0 }, values: { type: "array", minItems: 1, maxItems: 1000, items: { type: "array", minItems: 1, maxItems: 16384 } } }),
    operationObject(["type", "table", "enabled"], { type: { const: "setTableTotals" }, table: { type: "string", minLength: 1 }, enabled: { type: "boolean" }, columns: { type: "array", maxItems: 16384, items: operationObject(["column"], { column: { anyOf: [{ type: "string", minLength: 1 }, { type: "integer", minimum: 0 }] }, function: { type: "string", enum: ["average", "count", "countNums", "custom", "max", "min", "none", "stdDev", "sum", "var"] }, label: { type: "string" }, formula: { type: "string", minLength: 1 } }) } }),
    operationObject(["type", "sheet", "chartType", "sourceData"], { type: { const: "createChart" }, sheet: { type: "string", minLength: 1 }, chartType: { type: "string", enum: ["BarClustered", "ColumnClustered", "Line", "Pie"] }, sourceData: { type: "string", minLength: 1 }, seriesBy: { type: "string", enum: ["Auto", "Columns", "Rows"] }, name: { type: "string", minLength: 1 }, titleText: { type: "string" }, height: { type: "number", exclusiveMinimum: 0 }, width: { type: "number", exclusiveMinimum: 0 }, left: { type: "number", minimum: 0 }, top: { type: "number", minimum: 0 } }),
    operationObject(["type", "sheet", "chart"], { type: { const: "updateChart" }, sheet: { type: "string", minLength: 1 }, chart: { type: "string", minLength: 1 }, chartType: { type: "string", enum: ["BarClustered", "ColumnClustered", "Line", "Pie"] }, name: { type: "string", minLength: 1 }, titleText: { type: "string" }, height: { type: "number", exclusiveMinimum: 0 }, width: { type: "number", exclusiveMinimum: 0 }, left: { type: "number", minimum: 0 }, top: { type: "number", minimum: 0 }, sourceData: { type: "string", minLength: 1 }, seriesBy: { type: "string", enum: ["Auto", "Columns", "Rows"] } }),
    operationObject(["type", "sheet", "newName"], { type: { const: "renameSheet" }, sheet: { type: "string", minLength: 1 }, newName: { type: "string", minLength: 1, maxLength: 31 } }),
    operationObject(["type", "name", "formula"], { type: { const: "setDefinedName" }, name: { type: "string", minLength: 1 }, formula: { type: "string" } }),
    operationObject(["type"], { type: { const: "recalculate" } })
  ] }
};
const powerpointSelector = { slideIndex: { type: "integer", minimum: 0 }, shapeId: { type: ["string", "integer"] } };
const powerpointOperationsSchema = {
  type: "array", minItems: 1, maxItems: 100, items: { anyOf: [
    { ...replaceTextOperation, properties: { ...replaceTextOperation.properties, ...powerpointSelector } },
    operationObject(["type", "slideIndex", "shapeId", "text"], { type: { const: "setShapeText" }, ...powerpointSelector, text: { type: "string" } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "setShapeGeometry" }, ...powerpointSelector, x: { type: "integer" }, y: { type: "integer" }, width: { type: "integer", minimum: 0 }, height: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex", "shapeId", "rowIndex", "columnIndex", "text"], { type: { const: "setTableCell" }, ...powerpointSelector, rowIndex: { type: "integer", minimum: 0 }, columnIndex: { type: "integer", minimum: 0 }, text: { type: "string" } }),
    operationObject(["type", "slideIndex", "text", "x", "y", "width", "height"], { type: { const: "addTextBox" }, slideIndex: { type: "integer", minimum: 0 }, shapeId: { type: "integer", minimum: 1 }, name: { type: "string", minLength: 1 }, text: { type: "string" }, x: { type: "integer" }, y: { type: "integer" }, width: { type: "integer", minimum: 1 }, height: { type: "integer", minimum: 1 } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "deleteShape" }, ...powerpointSelector }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "setTextStyle" }, ...powerpointSelector, fontFamily: { type: "string", minLength: 1 }, fontSize: { type: "number", minimum: 1, maximum: 400 }, bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, color: { type: "string", pattern: "^[0-9A-Fa-f]{6}$" } }),
    operationObject(["type", "slideIndex", "shapeId", "base64", "contentType"], { type: { const: "replaceImage" }, ...powerpointSelector, base64: { type: "string", minLength: 1, maxLength: 36700160 }, contentType: { type: "string", enum: ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff"] } }),
    operationObject(["type", "slideIndex", "text"], { type: { const: "setNotes" }, slideIndex: { type: "integer", minimum: 0 }, text: { type: "string" } }),
    operationObject(["type", "slideIndex"], { type: { const: "duplicateSlide" }, slideIndex: { type: "integer", minimum: 0 }, toIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex"], { type: { const: "deleteSlide" }, slideIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex", "toIndex"], { type: { const: "moveSlide" }, slideIndex: { type: "integer", minimum: 0 }, toIndex: { type: "integer", minimum: 0 } })
  ] }
};
const pathTargetProperties = {
  path: { type: "string", description: "Item path relative to OneDrive root." },
  itemId: { type: "string", description: "Drive item ID." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const driveRecipientSchema = {
  type: "object",
  properties: {
    email: { type: "string", minLength: 1 },
    alias: { type: "string", minLength: 1 },
    objectId: { type: "string", minLength: 1 }
  },
  additionalProperties: false
};
const folderTargetProperties = {
  path: { type: "string", description: "Folder path relative to OneDrive root. Omit or use / for root." },
  itemId: { type: "string", description: "Drive item ID for the folder." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const destinationPresetProperties = {
  destinationParentPreset: presetSchema,
  destinationParentRelativePath: relativePathSchema
};
const parentPresetProperties = {
  parentPreset: presetSchema,
  parentRelativePath: relativePathSchema
};
const remotePresetProperties = {
  remotePreset: presetSchema,
  remoteRelativePath: {
    type: "string",
    description: "Destination path appended below the selected remotePreset, including filename."
  }
};
const itemTargetAnyOf = [
  { required: ["path"] },
  { required: ["itemId"] },
  { required: ["preset"] }
];
const remoteTargetAnyOf = [
  { required: ["remotePath"] },
  { required: ["remotePreset", "remoteRelativePath"] }
];

const tools = [
  {
    name: "onedrive_config",
    description: "Show OneDrive plugin configuration status without exposing secrets.",
    inputSchema: {
      type: "object",
      properties: {
        checkToken: {
          type: "boolean",
          default: false,
          description: "When true, try to get an access token from Keychain or memory."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_auth_device_start",
    description: "Start Microsoft device-code login for OneDrive and return the user code and verification URL.",
    inputSchema: {
      type: "object",
      properties: {
        tenant: { type: "string", description: "Optional tenant override. Defaults to configured tenant or common." },
        scopes: { type: "string", description: "Optional space-separated scope override." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_auth_device_poll",
    description: "Poll Microsoft token endpoint after the user completes device-code login, then store tokens in Keychain.",
    inputSchema: {
      type: "object",
      properties: {
        deviceCode: { type: "string", description: "Optional device_code from onedrive_auth_device_start. Defaults to the latest pending code." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_logout",
    description: "Forget cached OneDrive tokens from memory and optionally delete the Keychain token.",
    inputSchema: {
      type: "object",
      properties: {
        deleteKeychainToken: { type: "boolean", default: false },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation to delete the Keychain refresh token."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_me",
    description: "Return the signed-in Microsoft profile.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_drive",
    description: "Return metadata for the signed-in user's default OneDrive.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_doctor",
    description: "Run a bundled OneDrive plugin health check for config, auth, profile, drive, presets, and optional root listing.",
    inputSchema: {
      type: "object",
      properties: {
        checkRootList: {
          type: "boolean",
          default: true,
          description: "When true, list a few root items to verify Files.Read permissions."
        },
        rootListLimit: {
          type: "integer",
          minimum: 1,
          maximum: 20,
          default: 5
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_presets",
    description: "List configured OneDrive path presets and their root-relative paths.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_list",
    description: "List children in a OneDrive folder by path or item ID.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, default: 100 },
        select: { type: "string", description: "Optional Graph $select fields." },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_list_all",
    description: "List all children in a OneDrive folder by following Microsoft Graph pagination up to a safe item cap.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...folderTargetProperties,
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        select: { type: "string", description: "Optional Graph $select fields." },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_scan",
    description: "Recursively scan OneDrive folders from a starting folder, following pagination and subfolders up to safe caps.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        select: { type: "string", description: "Optional Graph $select fields." },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxResults: {
          type: "integer",
          minimum: 1,
          maximum: 5000,
          default: 500,
          description: "Maximum matching items returned in the response. The scan can inspect more items than it returns."
        },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        maxFolders: {
          type: "integer",
          minimum: 1,
          maximum: 10000,
          default: 1000,
          description: "Maximum folders to visit during recursive traversal."
        },
        nameContains: { type: "string", description: "Optional case-insensitive filename/folder-name substring filter." },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional file extension filter such as ['.pdf', 'pptx']. Folders are unaffected unless includeFolders is false."
        },
        includeFiles: { type: "boolean", default: true },
        includeFolders: { type: "boolean", default: true },
        stopAfterResults: {
          type: "boolean",
          default: false,
          description: "When true, stop traversal once maxResults matching items have been returned."
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_search",
    description: "Search the signed-in user's OneDrive.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        limit: { type: "integer", minimum: 1, default: 50 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_search_all",
    description: "Search OneDrive and follow Microsoft Graph pagination up to a safe item cap.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_find",
    description: "Cache-assisted remote-first file finder. Runs the canonical Graph query first, expands terms adaptively with bounded concurrency, ranks matches in memory, and optionally falls back to bounded recursive scans.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional hard extension filter such as ['.pdf', 'pptx']."
        },
        folderHints: {
          type: "array",
          items: { type: "string" },
          description: "Optional root-relative folders to try first during fallback scans."
        },
        includeFolders: { type: "boolean", default: false },
        maxResults: { type: "integer", minimum: 1, maximum: 50, default: 10 },
        maxSearchTerms: { type: "integer", minimum: 1, maximum: 12, default: 8 },
        searchConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Bounded concurrency for adaptive Graph search-term expansion after the canonical query."
        },
        searchPageSize: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        searchMaxItemsPerTerm: { type: "integer", minimum: 1, maximum: 500, default: 100 },
        scanFallback: { type: "boolean", default: true },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 10000, default: 1500 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 2000, default: 250 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 20 },
        scanConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Low-concurrency fallback scan fan-out across top folder hints."
        },
        minConfidenceForSearchOnly: {
          type: "integer",
          minimum: 0,
          maximum: 100,
          default: 78,
          description: "If the best search-only match scores below this, use scan fallback when enabled."
        },
        useCache: {
          type: "boolean",
          default: true,
          description: "Use the local metadata cache as a fast ranking source before live Graph search."
        },
        useContentIndex: {
          type: "boolean",
          default: true,
          description: "Use the local content index if it has already been built. This never fetches file contents during find."
        },
        contentMaxResults: {
          type: "integer",
          minimum: 0,
          maximum: 50,
          default: 10,
          description: "Maximum indexed-content matches to merge into ranked results."
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_find_all",
    description: "Broader cache-assisted remote-first file locator. Uses adaptive bounded-concurrency Graph term expansion, common-folder hints, and bounded scan fallback when needed.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional hard extension filter such as ['.pdf', 'pptx']."
        },
        folderHints: {
          type: "array",
          items: { type: "string" },
          description: "Optional root-relative folders to scan before the default common folders."
        },
        includeFolders: { type: "boolean", default: false },
        maxResults: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxSearchTerms: { type: "integer", minimum: 1, maximum: 12, default: 8 },
        searchConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Bounded concurrency for adaptive Graph search-term expansion after the canonical query."
        },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        scanConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Low-concurrency fallback scan fan-out across top folder hints."
        },
        searchPageSize: { type: "integer", minimum: 1, maximum: 200, default: 100 },
        searchMaxItemsPerTerm: { type: "integer", minimum: 1, maximum: 1000, default: 250 },
        useCache: {
          type: "boolean",
          default: true,
          description: "Use the local metadata cache as a fast ranking source before live Graph search."
        },
        useContentIndex: {
          type: "boolean",
          default: true,
          description: "Use the local content index if it has already been built. This never fetches file contents during find_all."
        },
        contentMaxResults: { type: "integer", minimum: 0, maximum: 100, default: 25 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_delta",
    description: "List recent changes from OneDrive delta sync, optionally continuing from a previous deltaLink or nextLink.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        deltaLink: { type: "string", description: "Previous @odata.deltaLink returned by onedrive_delta." },
        nextLink: { type: "string", description: "Previous @odata.nextLink returned by onedrive_delta." },
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_sync_status",
    description: "Show local metadata-cache status, delta cursor availability, and plugin storage locations.",
    inputSchema: {
      type: "object",
      properties: {
        includeSamples: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_cache_refresh",
    description: "Refresh the local metadata cache from a bounded recursive scan, or from delta when a previous cursor exists.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        mode: { type: "string", enum: ["auto", "scan", "delta"], default: "auto" },
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        includeFolders: { type: "boolean", default: true },
        replaceCache: { type: "boolean", default: false, description: "When true, clear existing metadata before the scan. Defaults to merge so bounded scans do not shrink a fuller cache." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_cache_clear",
    description: "Clear the local OneDrive metadata cache.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_content_index_refresh",
    description: "Build or refresh the optional local text content index from cached metadata or one selected file. This is the explicit expensive content-reading step.",
    inputSchema: {
      type: "object",
      properties: {
        ...pathTargetProperties,
        refreshMetadata: {
          type: "boolean",
          default: false,
          description: "When true, rebuild metadata with onedrive_cache_refresh before indexing cached files."
        },
        metadataMode: { type: "string", enum: ["auto", "delta", "scan"], default: "auto", description: "Metadata refresh strategy. Auto prefers the stored delta cursor and falls back to a bounded scan." },
        maxFiles: { type: "integer", minimum: 1, maximum: 1000, default: 100 },
        maxBytesPerFile: { type: "integer", minimum: 1024, maximum: 5242880, default: defaultMaxIndexedFileSize },
        concurrencyLimit: { type: "integer", minimum: 1, maximum: 8, default: 2 },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional extension allow-list. Defaults to configured supported indexed file types."
        },
        includeOfficeExport: {
          type: "boolean",
          default: false,
          description: "When true, try Microsoft Graph text export for Office-like files. Unsupported file types fail per item, not the whole refresh."
        },
        includeOfficeStructured: { type: "boolean", default: true, description: "Index structured Word, Excel, and PowerPoint Open XML content with semantic anchors." },
        maxOfficeSegments: { type: "integer", minimum: 1, maximum: 100000, default: 50000 },
        maxOfficePackageBytes: { type: "integer", minimum: 1024, maximum: 262144000, default: 52428800 },
        force: {
          type: "boolean",
          default: false,
          description: "Re-extract unchanged files instead of reusing index entries with matching ETag/cTag/mtime/size."
        },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_content_search",
    description: "Search the optional local text content index and return lightweight metadata plus snippets. Does not call Microsoft Graph.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        maxResults: { type: "integer", minimum: 1, maximum: 100, default: 20 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_index_refresh",
    description: "Incrementally refresh the structured local Office index using item identity and eTag/cTag freshness, preferring OneDrive delta metadata updates.",
    inputSchema: {
      type: "object",
      properties: {
        ...pathTargetProperties,
        refreshMetadata: { type: "boolean", default: true },
        metadataMode: { type: "string", enum: ["auto", "delta", "scan"], default: "auto" },
        maxFiles: { type: "integer", minimum: 1, maximum: 1000, default: 100 },
        concurrencyLimit: { type: "integer", minimum: 1, maximum: 8, default: 2 },
        force: { type: "boolean", default: false },
        maxOfficeSegments: { type: "integer", minimum: 1, maximum: 100000, default: 50000 },
        maxOfficePackageBytes: { type: "integer", minimum: 1024, maximum: 262144000, default: 52428800 },
        maxOfficeParagraphs: { type: "integer", minimum: 1, maximum: 10000, default: 10000 },
        maxOfficeCells: { type: "integer", minimum: 1, maximum: 50000, default: 50000 },
        maxOfficeSlides: { type: "integer", minimum: 1, maximum: 5000, default: 5000 },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_search",
    description: "Search the structured local Office index and return semantic paragraph, cell, formula, table, shape, slide-note, content-control, and comment anchors without Graph calls.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        maxResults: { type: "integer", minimum: 1, maximum: 100, default: 20 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_content_index_clear",
    description: "Clear the optional local OneDrive content index.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_office_capabilities",
    description: "Report available native Office editing backends, runtime health, account limitations, and supported Open XML formats.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_office_validate",
    description: "Download and safely validate a Word, Excel, or PowerPoint Open XML package without changing the remote file.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        expectedKind: { type: "string", enum: ["word", "excel", "powerpoint"] },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_word_get_document",
    description: "Read structured paragraphs, tables, content controls, headers, footers, comments, and package safety metadata from a .docx file.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        ...officeStructuredSearchProperties,
        maxParagraphs: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        includeHeadersFooters: { type: "boolean", default: true },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_excel_get_workbook",
    description: "Read structured worksheets, bounded cell/range selections, searchable values/formulas, table metadata, styles, and defined names from an .xlsx file using the local Open XML backend.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        ...officeStructuredSearchProperties,
        includeCells: { type: "boolean", default: true },
        includeTables: { type: "boolean", default: true },
        includeCharts: { type: "boolean", default: true },
        includePivots: { type: "boolean", default: true },
        includeFormulaDependencies: { type: "boolean", default: false },
        sheetNames: { type: "array", minItems: 1, maxItems: 100, uniqueItems: true, items: { type: "string", minLength: 1 }, description: "Optional worksheet-name selector." },
        address: { type: "string", pattern: "^[A-Za-z]{1,3}[1-9][0-9]*(?::[A-Za-z]{1,3}[1-9][0-9]*)?$", description: "Optional bounded A1 cell or range selector applied to selected worksheets." },
        maxCells: { type: "integer", minimum: 1, maximum: 50000, default: 5000 },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_powerpoint_get_presentation",
    description: "Read structured slides, shapes, text, placeholders, geometry, tables, images, notes, and package safety metadata from a .pptx file.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        ...officeStructuredSearchProperties,
        maxSlides: { type: "integer", minimum: 1, maximum: 5000, default: 500 },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_word_batch_update",
    description: "Preview or apply typed Word edits for text, paragraphs, styles, tables, content controls, external hyperlinks, and anchored comments. Documents containing tracked changes are refused; live commits require expected identity and the dry-run preview token.",
    inputSchema: {
      type: "object",
      required: ["operations"],
      anyOf: itemTargetAnyOf,
      properties: { ...officeBatchCommonProperties, operations: wordOperationsSchema, trackedChanges: { type: "string", enum: ["refuse"], default: "refuse", description: "Tracked-change markup is currently refused because native edits cannot safely preserve review semantics." } },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_excel_batch_update",
    description: "Preview or apply typed Excel cell/range, formatting, validation, pane, table-row, chart, sheet-name, defined-name, and recalculation edits. Auto uses Graph sessions only when every operation is supported for a business .xlsx file and Open XML otherwise.",
    inputSchema: {
      type: "object",
      required: ["operations"],
      anyOf: itemTargetAnyOf,
      properties: { ...officeBatchCommonProperties, operations: excelOperationsSchema, backend: { type: "string", enum: ["auto", "openxml", "graph"], default: "auto" } },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_powerpoint_batch_update",
    description: "Preview or apply typed PowerPoint text, text-box, text-style, shape deletion, geometry, image replacement, table, notes, duplicate, delete, and move-slide edits; live commits require a preview token.",
    inputSchema: {
      type: "object",
      required: ["operations"],
      anyOf: itemTargetAnyOf,
      properties: { ...officeBatchCommonProperties, operations: powerpointOperationsSchema },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_batch_transform",
    description: "Preflight and apply one atomic-plan Office transformation across up to 25 files. Every file is fully previewed before the first mutation; live partial state includes recovery backups and remaining items.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 25,
          items: {
            type: "object",
            required: ["kind", "operations"],
            anyOf: itemTargetAnyOf,
            properties: {
              ...officeTargetProperties,
              kind: { type: "string", enum: ["word", "excel", "powerpoint"] },
              operations: { type: "array", minItems: 1, maxItems: 100, items: { type: "object" } },
              expectedName: { type: "string" },
              expectedId: { type: "string" },
              backend: { type: "string", enum: ["auto", "openxml", "graph"], default: "auto" },
              allowMacros: { type: "boolean", default: false }
            },
            additionalProperties: false
          }
        },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        previewToken: previewTokenSchema,
        createBackup: { type: "boolean", default: true },
        verify: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_backups",
    description: "List plugin-managed Office backups with stable backup IDs and original remote item metadata.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", minLength: 1 },
        kind: { type: "string", enum: ["word", "excel", "powerpoint"] },
        limit: { type: "integer", minimum: 1, maximum: 500, default: 100 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_compare_backup",
    description: "Compare one plugin-managed Office backup with the current content of its original stable remote item ID.",
    inputSchema: {
      type: "object",
      required: ["backupId"],
      properties: { backupId: { type: "string", pattern: "^[0-9a-fA-F-]{36}$" }, maxChanges: { type: "integer", minimum: 1, maximum: 500, default: 100 } },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_restore_backup",
    description: "Preview or restore an Office backup to its original stable item ID. Live restore requires expected ID/eTag and the dry-run preview token.",
    inputSchema: {
      type: "object",
      required: ["backupId"],
      properties: {
        backupId: { type: "string", pattern: "^[0-9a-fA-F-]{36}$" },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        expectedId: { type: "string" },
        expectedETag: { type: "string" },
        previewToken: previewTokenSchema,
        verify: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_get_info",
    description: "Get metadata for a OneDrive item by path or item ID.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        includeDeletedItems: {
          type: "boolean",
          default: false,
          description: "When true and itemId is used, ask Graph to include deleted items. Supported by Microsoft for OneDrive Personal."
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_read_text",
    description: "Read a bounded text file from OneDrive by path or item ID.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        maxBytes: { type: "integer", minimum: 1, maximum: 10485760, default: textFileLimit },
        force: {
          type: "boolean",
          default: false,
          description: "When true, bypass MIME/extension text checks while keeping maxBytes enforcement."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_preview",
    description: "Return a bounded safe preview for text files and Graph-supported document text exports.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        maxBytes: { type: "integer", minimum: 1, maximum: 1048576, default: 65536 },
        preferExportText: {
          type: "boolean",
          default: true,
          description: "For likely Office files, ask Graph for a text export before falling back to raw text checks."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download",
    description: "Download a OneDrive file to a local path.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/<filename>." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download_excel",
    description: "Download an Excel-like file (.xlsx, .xls, .xlsm, .xlsb, .csv, .ods) with extension/MIME checks.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/excel/<filename>." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download_word",
    description: "Download a Word-like document (.docx, .doc, .docm, .rtf, .odt) with extension/MIME checks.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/word/<filename>." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download_powerpoint",
    description: "Download a PowerPoint-like file (.pptx, .ppt, .pptm, .ppsx, .odp) with extension/MIME checks.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/powerpoint/<filename>." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_export_pdf",
    description: "Ask Microsoft Graph to export a supported OneDrive document to PDF and save it locally.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/export/<name>.pdf." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_export_text",
    description: "Ask Microsoft Graph to export a supported OneDrive document to plain text and save it locally when Graph supports that conversion.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/export/<name>.txt." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_upload",
    description: "Upload a local file to OneDrive. Uses simple upload for small files and upload sessions for large files.",
    inputSchema: {
      type: "object",
      required: ["localPath"],
      anyOf: remoteTargetAnyOf,
      properties: {
        localPath: { type: "string", description: "Absolute or workspace-relative local file path." },
        remotePath: { type: "string", description: "Destination path relative to OneDrive root, including filename." },
        ...remotePresetProperties,
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" },
        uploadMode: { type: "string", enum: ["auto", "simple", "session"], default: "auto" },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to read from a locally synced OneDrive folder. Prefer remote plugin operations instead."
        },
        chunkSize: {
          type: "integer",
          minimum: 327680,
          maximum: 62914560,
          default: defaultUploadChunkSize,
          description: "Upload-session chunk size. Must be a multiple of 320 KiB."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_write_text",
    description: "Create or replace a small text file in OneDrive.",
    inputSchema: {
      type: "object",
      required: ["content"],
      anyOf: remoteTargetAnyOf,
      properties: {
        remotePath: { type: "string", description: "Destination path relative to OneDrive root, including filename." },
        ...remotePresetProperties,
        content: { type: "string" },
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_create_folder",
    description: "Create a folder in OneDrive.",
    inputSchema: {
      type: "object",
      required: ["name"],
      properties: {
        parentPath: { type: "string", description: "Parent folder path relative to OneDrive root. Defaults to root." },
        parentItemId: { type: "string", description: "Parent folder drive item ID." },
        ...parentPresetProperties,
        name: { type: "string", minLength: 1 },
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_rename",
    description: "Rename a OneDrive item by path or item ID.",
    inputSchema: {
      type: "object",
      required: ["newName"],
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        newName: { type: "string", minLength: 1 },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this renames a file or folder."
        },
        expectedName: { type: "string", description: "Optional safety check: item name must match before renaming." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match before renaming." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_move",
    description: "Move a OneDrive item to another folder, optionally renaming it. Refuses to move the OneDrive root.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        path: { type: "string", description: "Source item path relative to OneDrive root." },
        itemId: { type: "string", description: "Source drive item ID." },
        preset: presetSchema,
        relativePath: relativePathSchema,
        destinationParentPath: { type: "string", description: "Destination folder path relative to root. Omit or use / for root." },
        destinationParentItemId: { type: "string", description: "Destination folder drive item ID." },
        ...destinationPresetProperties,
        newName: { type: "string", description: "Optional new name after moving." },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this moves a file or folder."
        },
        expectedName: { type: "string", description: "Optional safety check: source item name must match." },
        expectedId: { type: "string", description: "Optional safety check: source item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_copy",
    description: "Start a OneDrive copy operation to another folder. Copy is asynchronous; optionally wait for completion.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        path: { type: "string", description: "Source item path relative to OneDrive root." },
        itemId: { type: "string", description: "Source drive item ID." },
        preset: presetSchema,
        relativePath: relativePathSchema,
        destinationParentPath: { type: "string", description: "Destination folder path relative to root. Omit or use / for root." },
        destinationParentItemId: { type: "string", description: "Destination folder drive item ID." },
        ...destinationPresetProperties,
        newName: { type: "string", description: "Optional copied item name." },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this copies a file or folder."
        },
        waitForCompletion: { type: "boolean", default: false },
        timeoutSeconds: { type: "integer", minimum: 1, maximum: 300, default: 60 },
        expectedName: { type: "string", description: "Optional safety check: source item name must match." },
        expectedId: { type: "string", description: "Optional safety check: source item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_create_sharing_link",
    description: "Create a OneDrive sharing link. Defaults to dry-run and requires confirmed=true plus dryRun=false to change sharing.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        type: { type: "string", enum: ["view", "edit", "embed"], default: "view" },
        scope: { type: "string", enum: ["anonymous", "organization", "users"], default: "anonymous" },
        password: { type: "string", minLength: 1, description: "Optional sharing-link password. Never returned in dry-runs or audit logs." },
        expirationDateTime: { type: "string", minLength: 1, description: "Optional ISO 8601 link expiration timestamp." },
        retainInheritedPermissions: { type: "boolean" },
        includePermissionDiff: {
          type: "boolean",
          default: true,
          description: "When true, include before-permissions in dry-run and before/after permission diff for live creation."
        },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this can expose access to a file or folder."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_invite_permission",
    description: "Invite users or groups to a OneDrive item. Defaults to dry-run and silent direct grants unless sendInvitation=true is set.",
    inputSchema: {
      type: "object",
      required: ["recipients"],
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        recipients: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: driveRecipientSchema,
          description: "Recipients to grant access to. Each item must include exactly one of email, alias, or objectId."
        },
        role: { type: "string", enum: ["read", "write"], default: "read" },
        sendInvitation: { type: "boolean", default: false },
        requireSignIn: { type: "boolean", default: true },
        message: { type: "string", description: "Optional invitation message. Only useful when sendInvitation=true; never written to audit logs." },
        password: { type: "string", minLength: 1, description: "Optional invite password when supported by Microsoft Graph. Never returned in dry-runs or audit logs." },
        expirationDateTime: { type: "string", minLength: 1, description: "Optional ISO 8601 expiration timestamp." },
        retainInheritedPermissions: { type: "boolean" },
        includePermissionDiff: {
          type: "boolean",
          default: true,
          description: "When true, include before-permissions in dry-run and before/after permission diff for live invitation."
        },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this grants access to a file or folder."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_revoke_permission",
    description: "Preview or revoke a OneDrive sharing permission. Defaults to dry-run and requires confirmation plus expected item identity for live revocation.",
    inputSchema: {
      type: "object",
      required: ["permissionId"],
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        permissionId: { type: "string", minLength: 1 },
        includePermissions: {
          type: "boolean",
          default: true,
          description: "When true, include before permissions in dry-run and before/after permission diff for live revocation."
        },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this removes sharing access."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_revoke_permissions",
    description: "Preview or revoke multiple OneDrive sharing permissions. Live revoke preflights every item before any DELETE.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            required: ["permissionId"],
            anyOf: itemTargetAnyOf,
            properties: {
              ...pathTargetProperties,
              permissionId: { type: "string", minLength: 1 },
              expectedName: { type: "string" },
              expectedId: { type: "string" }
            },
            additionalProperties: false
          }
        },
        includePermissions: { type: "boolean", default: true },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation for live batch permission revocation."
        },
        previewToken: previewTokenSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_restore_deleted",
    description: "Restore a deleted OneDrive item by item ID. Defaults to dry-run and may require Files.ReadWrite.All on personal OneDrive.",
    inputSchema: {
      type: "object",
      required: ["itemId"],
      properties: {
        itemId: { type: "string", description: "Deleted drive item ID to restore." },
        destinationParentPath: { type: "string", description: "Optional restore destination folder path relative to root." },
        destinationParentItemId: { type: "string", description: "Optional restore destination folder drive item ID." },
        ...destinationPresetProperties,
        newName: { type: "string", description: "Optional new name when restoring." },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this restores a deleted item."
        },
        previewToken: previewTokenSchema,
        expectedId: { type: "string", description: "Optional safety check: itemId must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_permissions",
    description: "List sharing and permission grants on a OneDrive file or folder.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_get_info",
    description: "Get metadata for up to 20 OneDrive items using Microsoft Graph batching.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 20,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: pathTargetProperties,
            additionalProperties: false
          }
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_permissions",
    description: "Audit sharing permissions for up to 20 OneDrive files or folders using Microsoft Graph batching.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 20,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: pathTargetProperties,
            additionalProperties: false
          }
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_download",
    description: "Download multiple OneDrive files serially with one result per item.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: {
              ...pathTargetProperties,
              localPath: { type: "string" },
              overwrite: { type: "boolean", default: false }
            },
            additionalProperties: false
          }
        },
        destinationFolder: { type: "string", description: "Optional local folder for downloaded files." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_delete",
    description: "Preview or delete multiple OneDrive items. Live deletes require dryRun=false and confirmed=true.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: {
              ...pathTargetProperties,
              expectedName: { type: "string" },
              expectedId: { type: "string" }
            },
            additionalProperties: false
          }
        },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        previewToken: previewTokenSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_move",
    description: "Preview or move multiple OneDrive items to one destination folder.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: {
              path: { type: "string", description: "Source item path relative to OneDrive root." },
              itemId: { type: "string", description: "Source drive item ID." },
              preset: presetSchema,
              relativePath: relativePathSchema,
              newName: { type: "string" },
              expectedName: { type: "string" },
              expectedId: { type: "string" }
            },
            additionalProperties: false
          }
        },
        destinationParentPath: { type: "string", description: "Destination folder path relative to root. Omit or use / for root." },
        destinationParentItemId: { type: "string", description: "Destination folder drive item ID." },
        ...destinationPresetProperties,
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation for live batch moves."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_update_file",
    description: "Safe file update workflow: checkout downloads a remote file with a manifest; commit uploads an edited file back after conflict checks.",
    inputSchema: {
      type: "object",
      required: ["mode", "remotePath"],
      properties: {
        mode: { type: "string", enum: ["checkout", "commit"] },
        remotePath: { type: "string", description: "Remote OneDrive path including filename." },
        itemId: { type: "string", description: "Optional item ID for checkout." },
        localPath: { type: "string", description: "Checkout destination or commit source. Checkout defaults to ~/.codex/onedrive-plugin/updates." },
        manifestPath: { type: "string", description: "Optional checkout manifest path. Defaults to localPath + .onedrive-update.json." },
        conflictCheck: { type: "boolean", default: true },
        force: { type: "boolean", default: false },
        createBackup: { type: "boolean", default: true },
        verify: { type: "boolean", default: true },
        overwriteLocal: { type: "boolean", default: false },
        overwriteManifest: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_recent",
    description: "List recently used OneDrive files from Microsoft Graph.",
    inputSchema: {
      type: "object",
      properties: {
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_large_files",
    description: "Find large files by scanning OneDrive and sorting by size.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        minBytes: { type: "integer", minimum: 0, default: 104857600 },
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_duplicates",
    description: "Find likely duplicate files by scanning OneDrive and grouping by hash when available, otherwise by name and size.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_shared_by_me",
    description: "Scan files and folders and return items with explicit sharing permissions.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 1000, default: 250 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 15 },
        includeFolders: { type: "boolean", default: true },
        includeOwnerPermissions: { type: "boolean", default: false, description: "When true, include owner/self grants in addition to external sharing permissions." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_public_links",
    description: "Scan files and folders and return anonymous sharing links.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 1000, default: 250 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 15 },
        includeFolders: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_audit_recent",
    description: "Read recent local OneDrive live-mutation audit log entries.",
    inputSchema: {
      type: "object",
      properties: {
        limit: { type: "integer", minimum: 1, maximum: 500, default: 50 },
        tool: { type: "string", description: "Optional exact tool name filter, such as onedrive_upload." },
        status: { type: "string", description: "Optional mutation status filter, such as success or failed." },
        pathContains: { type: "string", description: "Optional case-insensitive filter matched against target, before, and after paths/names." },
        since: { type: "string", description: "Optional ISO timestamp; only entries at or after this time are returned." },
        until: { type: "string", description: "Optional ISO timestamp; only entries at or before this time are returned." },
        newestFirst: { type: "boolean", default: true, description: "Return newest entries first. Set false to preserve log order." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_audit_export",
    description: "Export the local OneDrive live-mutation audit log to a JSONL file.",
    inputSchema: {
      type: "object",
      properties: {
        localPath: { type: "string", description: "Optional export path. Defaults to ~/.codex/onedrive-plugin/audit/export-<timestamp>.jsonl." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_audit_clear",
    description: "Clear the local OneDrive live-mutation audit log. Requires confirmed=true.",
    inputSchema: {
      type: "object",
      properties: {
        confirmed: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_delete",
    description: "Delete a OneDrive item by path or item ID. Defaults to dry-run.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this deletes a file or folder."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match before deleting." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match before deleting." }
      },
      additionalProperties: false
    }
  }
];

const toolByName = new Map(tools.map((tool) => [tool.name, tool]));

function schemaTypeMatches(value, type) {
  if (Array.isArray(type)) return type.some((candidate) => candidate === "null" ? value === null : schemaTypeMatches(value, candidate));
  if (type === "array") return Array.isArray(value);
  if (type === "integer") return Number.isInteger(value);
  if (type === "number") return typeof value === "number" && Number.isFinite(value);
  if (type === "object") return value !== null && typeof value === "object" && !Array.isArray(value);
  return typeof value === type;
}

function validationDetail(path, message) {
  return { path, message };
}

function cloneDefault(value) {
  if (value === null || typeof value !== "object") return value;
  return JSON.parse(JSON.stringify(value));
}

function validateSchemaValue(value, schema = {}, path = "$") {
  const details = [];
  let normalized = value;

  if (schema.default !== undefined && normalized === undefined) {
    normalized = cloneDefault(schema.default);
  }

  if (normalized === undefined) return { ok: true, value: normalized, details };

  if (schema.anyOf?.length && !schema.type) {
    const branches = schema.anyOf.map((branch) => validateSchemaValue(normalized, branch, path));
    const matched = branches.filter((branch) => branch.ok);
    if (!matched.length) {
      details.push(validationDetail(path, "Value does not match any allowed schema."));
      const mostSpecific = branches.sort((left, right) => left.details.length - right.details.length)[0];
      if (mostSpecific) details.push(...mostSpecific.details);
      return { ok: false, value: normalized, details };
    }
    return { ok: true, value: matched[0].value, details };
  }

  if (schema.type && !schemaTypeMatches(normalized, schema.type)) {
    details.push(validationDetail(path, `Expected ${Array.isArray(schema.type) ? schema.type.join(" or ") : schema.type}.`));
    return { ok: false, value: normalized, details };
  }

  if (schema.const !== undefined && normalized !== schema.const) {
    details.push(validationDetail(path, `Expected constant value ${JSON.stringify(schema.const)}.`));
  }

  if (schema.enum && !schema.enum.includes(normalized)) {
    details.push(validationDetail(path, `Expected one of: ${schema.enum.join(", ")}.`));
  }

  if ((schema.type === "number" || schema.type === "integer") && typeof normalized === "number") {
    if (schema.minimum !== undefined && normalized < schema.minimum) {
      details.push(validationDetail(path, `Must be >= ${schema.minimum}.`));
    }
    if (schema.maximum !== undefined && normalized > schema.maximum) {
      details.push(validationDetail(path, `Must be <= ${schema.maximum}.`));
    }
    if (schema.exclusiveMinimum !== undefined && normalized <= schema.exclusiveMinimum) {
      details.push(validationDetail(path, `Must be > ${schema.exclusiveMinimum}.`));
    }
    if (schema.exclusiveMaximum !== undefined && normalized >= schema.exclusiveMaximum) {
      details.push(validationDetail(path, `Must be < ${schema.exclusiveMaximum}.`));
    }
  }

  if (schema.type === "string" && typeof normalized === "string") {
    if (schema.minLength !== undefined && normalized.length < schema.minLength) {
      details.push(validationDetail(path, `Must be at least ${schema.minLength} characters.`));
    }
    if (schema.maxLength !== undefined && normalized.length > schema.maxLength) {
      details.push(validationDetail(path, `Must be at most ${schema.maxLength} characters.`));
    }
    if (schema.pattern !== undefined && !(new RegExp(schema.pattern).test(normalized))) {
      details.push(validationDetail(path, `Must match pattern ${schema.pattern}.`));
    }
  }

  if (schema.type === "array" && Array.isArray(normalized)) {
    if (schema.minItems !== undefined && normalized.length < schema.minItems) {
      details.push(validationDetail(path, `Must contain at least ${schema.minItems} item(s).`));
    }
    if (schema.maxItems !== undefined && normalized.length > schema.maxItems) {
      details.push(validationDetail(path, `Must contain at most ${schema.maxItems} item(s).`));
    }
    if (schema.items) {
      normalized = normalized.map((item, index) => {
        const child = validateSchemaValue(item, schema.items, `${path}[${index}]`);
        details.push(...child.details);
        return child.value;
      });
    }
  }

  if (schema.type === "object" && normalized && typeof normalized === "object" && !Array.isArray(normalized)) {
    const properties = schema.properties || {};
    normalized = { ...normalized };
    for (const required of schema.required || []) {
      if (!Object.hasOwn(normalized, required)) {
        details.push(validationDetail(`${path}.${required}`, "Required field is missing."));
      }
    }
    if (schema.additionalProperties === false) {
      for (const key of Object.keys(normalized)) {
        if (!Object.hasOwn(properties, key)) {
          details.push(validationDetail(`${path}.${key}`, "Unknown property."));
        }
      }
    }
    for (const [key, propertySchema] of Object.entries(properties)) {
      const child = validateSchemaValue(normalized[key], propertySchema, `${path}.${key}`);
      details.push(...child.details);
      if (child.value !== undefined || Object.hasOwn(normalized, key)) normalized[key] = child.value;
    }
    if (schema.anyOf?.length) {
      const branchLabels = schema.anyOf
        .map((branch) => (branch.required || []).join(" + "))
        .filter(Boolean);
      const branchMatches = schema.anyOf.filter((branch) =>
        (branch.required || []).every((key) => Object.hasOwn(normalized, key))
      );
      if (branchMatches.length === 0) {
        details.push(validationDetail(path, `Must include one target option: ${branchLabels.join(" or ")}.`));
      } else if (branchMatches.length > 1) {
        details.push(validationDetail(path, `Must include exactly one target option: ${branchLabels.join(" or ")}.`));
      }
    }
  }

  return { ok: details.length === 0, value: normalized, details };
}

function validateToolArguments(name, args = {}) {
  const tool = toolByName.get(name);
  if (!tool) return { ok: true, args };
  const result = validateSchemaValue(args || {}, tool.inputSchema || { type: "object" }, "$");
  if (result.ok) return { ok: true, args: result.value };
  return {
    ok: false,
    error: {
      error: "invalid_arguments",
      tool: name,
      details: result.details
    }
  };
}

function readLocalConfig() {
  try {
    return JSON.parse(readFileSync(configPath, "utf8"));
  } catch {
    return {};
  }
}

function config(overrides = {}) {
  const clientId = overrides.clientId || process.env.ONEDRIVE_CLIENT_ID || localConfig.clientId;
  const tenant = overrides.tenant || process.env.ONEDRIVE_TENANT || localConfig.tenant || "common";
  const scopes = overrides.scopes || process.env.ONEDRIVE_SCOPES || localConfig.scopes || "offline_access User.Read Files.ReadWrite";
  const keychainService = process.env.ONEDRIVE_KEYCHAIN_SERVICE || localConfig.keychainService || "Codex OneDrive";
  return { clientId, tenant, scopes, keychainService };
}

function boolSetting(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  return ["1", "true", "yes", "on"].includes(String(value).toLowerCase());
}

function intSetting(value, defaultValue, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return defaultValue;
  return Math.min(Math.max(Math.trunc(number), min), max);
}

function listSetting(value, defaultValue) {
  if (Array.isArray(value)) return value;
  if (typeof value === "string") return value.split(/[,\s]+/).filter(Boolean);
  return defaultValue;
}

function pluginSettings() {
  const cfg = localConfig.settings || {};
  const indexing = localConfig.indexing || {};
  const supportedIndexedFileTypes = normalizeExtensions(listSetting(
    process.env.ONEDRIVE_INDEX_EXTENSIONS || indexing.supportedExtensions || cfg.supportedIndexedFileTypes,
    [".txt", ".md", ".csv", ".json", ".jsonl", ".xml", ".yaml", ".yml", ".html", ".css", ".js", ".mjs", ".ts", ".tsx", ".py", ".sql", ".log"]
  ));
  return {
    storageRoot,
    cacheRoot,
    cachePath,
    contentIndexPath,
    maxScanDepth: intSetting(process.env.ONEDRIVE_MAX_SCAN_DEPTH ?? cfg.maxScanDepth, 25, 0, 50),
    maxIndexedFileSize: intSetting(process.env.ONEDRIVE_MAX_INDEXED_FILE_SIZE ?? indexing.maxFileSize ?? cfg.maxIndexedFileSize, defaultMaxIndexedFileSize, 1024, textFileLimit),
    supportedIndexedFileTypes: [...supportedIndexedFileTypes],
    cacheTtlSeconds: intSetting(process.env.ONEDRIVE_CACHE_TTL_SECONDS ?? cfg.cacheTtlSeconds, 900, 0, 30 * 24 * 60 * 60),
    concurrencyLimit: intSetting(process.env.ONEDRIVE_CONCURRENCY_LIMIT ?? cfg.concurrencyLimit, 2, 1, 8),
    deltaSyncEnabled: boolSetting(process.env.ONEDRIVE_DELTA_SYNC_ENABLED ?? cfg.deltaSyncEnabled, true),
    contentIndexEnabled: boolSetting(process.env.ONEDRIVE_CONTENT_INDEX_ENABLED ?? indexing.enabled ?? cfg.contentIndexEnabled, true),
    includeOfficeTextExport: boolSetting(process.env.ONEDRIVE_INDEX_OFFICE_EXPORT ?? indexing.includeOfficeTextExport, false),
    fullReindexCommand: "onedrive_content_index_refresh({ force: true, refreshMetadata: true })"
  };
}

function pathPresets() {
  return {
    documents: "Documents",
    desktop: "Desktop",
    pictures: "Pictures",
    screenshots: "Pictures/Screenshots",
    ...(localConfig.pathPresets || {})
  };
}

function requireClientId(cfg = config()) {
  if (!cfg.clientId) {
    throw new Error("ONEDRIVE_CLIENT_ID is not configured. Run scripts/configure.zsh from the OneDrive plugin directory, or set ONEDRIVE_CLIENT_ID.");
  }
}

function tokenUrl(tenant) {
  return `${identityBaseUrl()}/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
}

function deviceCodeUrl(tenant) {
  return `${identityBaseUrl()}/${encodeURIComponent(tenant)}/oauth2/v2.0/devicecode`;
}

function identityBaseUrl() {
  return (process.env.ONEDRIVE_IDENTITY_BASE_URL || "https://login.microsoftonline.com").replace(/\/+$/, "");
}

function isMsaOnlyEndpointError(error) {
  const message = `${error?.body?.error || ""} ${error?.body?.error_description || ""} ${error?.message || ""}`;
  return message.includes("AADSTS9002331") || message.includes("/consumers endpoint");
}

async function postFormWithConsumerFallback(kind, cfg, params) {
  try {
    return { body: await postForm(kind === "device" ? deviceCodeUrl(cfg.tenant) : tokenUrl(cfg.tenant), params), tenant: cfg.tenant };
  } catch (error) {
    if (cfg.tenant === "common" && isMsaOnlyEndpointError(error)) {
      return { body: await postForm(kind === "device" ? deviceCodeUrl("consumers") : tokenUrl("consumers"), params), tenant: "consumers" };
    }
    throw error;
  }
}

function keychainAccount() {
  return "tokens";
}

function getKeychainToken(cfg = config()) {
  try {
    const raw = execFileSync("security", [
      "find-generic-password",
      "-a", keychainAccount(),
      "-s", cfg.keychainService,
      "-w"
    ], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function setKeychainToken(token, cfg = config()) {
  const payload = JSON.stringify(token);
  try {
    execFileSync("security", [
      "add-generic-password",
      "-U",
      "-a", keychainAccount(),
      "-s", cfg.keychainService,
      "-w", payload
    ], { stdio: "ignore" });
  } catch (error) {
    throw new Error(`Could not store OneDrive token in Keychain: ${error.message}`);
  }
}

function deleteKeychainToken(cfg = config()) {
  try {
    execFileSync("security", [
      "delete-generic-password",
      "-a", keychainAccount(),
      "-s", cfg.keychainService
    ], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

function publicConfig() {
  const cfg = config();
  const stored = getKeychainToken(cfg);
  return {
    clientIdConfigured: Boolean(cfg.clientId),
    tenant: cfg.tenant,
    scopes: cfg.scopes,
    keychainService: cfg.keychainService,
    keychainTokenConfigured: Boolean(stored?.refresh_token),
    configPath,
    pathPresets: pathPresets(),
    settings: pluginSettings()
  };
}

function emptyMetadataCache() {
  return {
    version: 3,
    createdAt: new Date().toISOString(),
    updatedAt: null,
    deltaLink: null,
    deltaNextLink: null,
    deltaTarget: null,
    scanRoot: null,
    pathRootsById: {},
    itemCount: 0,
    itemsById: {},
    pathsByLower: {}
  };
}

async function ensurePrivateDirectory(path) {
  await mkdir(path, { recursive: true, mode: 0o700 });
  await chmod(path, 0o700);
}

async function hardenPrivateFile(path) {
  try {
    await chmod(path, 0o600);
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
}

async function writePrivateFile(path, data, options = {}) {
  await writeFile(path, data, { ...options, mode: 0o600 });
  await hardenPrivateFile(path);
}

async function writePrivateFileAtomic(path, data, options = {}) {
  await ensurePrivateDirectory(dirname(path));
  const temporaryPath = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
  let handle;
  try {
    handle = await open(temporaryPath, "wx", 0o600);
    await handle.writeFile(data, options);
    await handle.sync();
    await handle.close();
    handle = null;
    await renameFile(temporaryPath, path);
    await hardenPrivateFile(path);
  } finally {
    await handle?.close().catch(() => null);
    await rm(temporaryPath, { force: true }).catch(() => null);
  }
}

function collisionPath(path, suffix) {
  const parsed = parse(path);
  return join(parsed.dir, `${parsed.name} (${suffix})${parsed.ext}`);
}

async function reserveLocalDestination(preferredPath, options = {}) {
  const preferred = resolve(preferredPath);
  await ensurePrivateDirectory(dirname(preferred));
  if (options.overwrite === true) return { path: preferred, reserved: false };
  const allowAlternate = options.allowAlternate === true;
  for (let suffix = 1; suffix <= 10_000; suffix += 1) {
    const candidate = suffix === 1 ? preferred : collisionPath(preferred, suffix);
    try {
      const handle = await open(candidate, "wx", 0o600);
      await handle.close();
      return { path: candidate, reserved: true };
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
      if (!allowAlternate) {
        throw new Error(`Local file already exists: ${preferred}. Pass overwrite: true to replace it.`);
      }
    }
  }
  throw new Error(`Could not reserve a unique local destination near ${preferred}.`);
}

async function localFileVersion(path) {
  try {
    const fileStat = await stat(path);
    return `${fileStat.dev}:${fileStat.ino}:${fileStat.size}:${fileStat.mtimeMs}`;
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  }
}

async function withFileLock(lockPath, fn, options = {}) {
  const timeoutMs = options.timeoutMs ?? 10_000;
  const staleMs = options.staleMs ?? 60_000;
  const startedAt = Date.now();
  await ensurePrivateDirectory(dirname(lockPath));
  let handle;
  while (!handle) {
    try {
      handle = await open(lockPath, "wx", 0o600);
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
      try {
        const lockStat = await stat(lockPath);
        if (Date.now() - lockStat.mtimeMs > staleMs) {
          await rm(lockPath, { force: true });
          continue;
        }
      } catch (statError) {
        if (statError.code !== "ENOENT") throw statError;
        continue;
      }
      if (Date.now() - startedAt >= timeoutMs) throw new Error(`Timed out waiting for local storage lock: ${lockPath}`);
      await sleep(Math.min(100, 10 + Math.floor((Date.now() - startedAt) / 10)));
    }
  }
  try {
    return await fn();
  } finally {
    await handle.close().catch(() => null);
    await rm(lockPath, { force: true }).catch(() => null);
  }
}

async function runSerialized(queueName, fn) {
  const previous = queueName === "metadata" ? metadataMutationQueue : contentIndexMutationQueue;
  let release;
  const current = new Promise((resolve) => { release = resolve; });
  if (queueName === "metadata") metadataMutationQueue = previous.catch(() => null).then(() => current);
  else contentIndexMutationQueue = previous.catch(() => null).then(() => current);
  await previous.catch(() => null);
  try {
    return await fn();
  } finally {
    release();
  }
}

function normalizeMetadataCache(parsed = {}) {
  const cache = {
    ...emptyMetadataCache(),
    ...parsed,
    version: 3,
    itemsById: parsed.itemsById || {},
    pathsByLower: parsed.pathsByLower || {},
    pathRootsById: parsed.pathRootsById || {}
  };
  if ((cache.deltaLink || cache.deltaNextLink) && !parsed.deltaTarget) {
    cache.deltaLink = null;
    cache.deltaNextLink = null;
    cache.deltaTarget = null;
  }
  return cache;
}

async function readMetadataCacheFromDisk() {
  await ensurePrivateDirectory(cacheRoot);
  try {
    const parsed = JSON.parse(await readFile(cachePath, "utf8"));
    await hardenPrivateFile(cachePath);
    return normalizeMetadataCache(parsed);
  } catch (error) {
    if (error?.code === "ENOENT") return emptyMetadataCache();
    throw error;
  }
}

async function loadMetadataCache() {
  if (!metadataCacheLoadPromise) {
    metadataCacheLoadPromise = (async () => {
      try {
        const diskVersion = await localFileVersion(cachePath);
        if (metadataCacheMemory && diskVersion === metadataCacheFileVersion) return metadataCacheMemory;
        await withFileLock(metadataCacheLockPath, async () => {
          const lockedVersion = await localFileVersion(cachePath);
          if (metadataCacheMemory && lockedVersion === metadataCacheFileVersion) return;
          metadataCacheMemory = await readMetadataCacheFromDisk();
          metadataCacheFileVersion = await localFileVersion(cachePath);
        });
      } catch (error) {
        recordLocalWarning("metadata cache read", error);
        metadataCacheMemory = emptyMetadataCache();
        metadataCacheFileVersion = null;
      } finally {
        metadataCacheLoadPromise = null;
      }
      return metadataCacheMemory;
    })();
  }
  return await metadataCacheLoadPromise;
}

async function saveMetadataCache(cache) {
  cache.updatedAt = new Date().toISOString();
  cache.itemCount = Object.keys(cache.itemsById || {}).length;
  await ensurePrivateDirectory(cacheRoot);
  await writePrivateFileAtomic(cachePath, JSON.stringify(cache));
  metadataCacheMemory = cache;
  metadataCacheFileVersion = await localFileVersion(cachePath);
  const store = toolCallContext.getStore();
  if (store) store.metadataCacheWrites = (store.metadataCacheWrites || 0) + 1;
  return cache;
}

async function withMetadataCacheBatch(fn) {
  return await fn();
}

function cachePathKey(remotePath = "") {
  return cleanPath(remotePath).toLowerCase();
}

function mergeDefinedMetadata(previous = {}, incoming = {}) {
  const merged = { ...(previous || {}) };
  for (const [key, value] of Object.entries(incoming || {})) {
    if (value === undefined) continue;
    if ((key === "file" || key === "folder") && value && typeof value === "object") {
      merged[key] = mergeDefinedMetadata(previous?.[key] || {}, value);
    } else {
      merged[key] = value;
    }
  }
  return merged;
}

function remoteParentPath(remotePath = "") {
  const clean = cleanPath(remotePath);
  const separator = clean.lastIndexOf("/");
  return separator < 0 ? "" : clean.slice(0, separator);
}

function cachedParentRemotePath(cache, parentId) {
  if (!parentId) return null;
  const parent = cache.itemsById?.[parentId];
  if (parent?.remotePath !== undefined) return cleanPath(parent.remotePath);
  if (Object.hasOwn(cache.pathRootsById || {}, parentId)) return cleanPath(cache.pathRootsById[parentId] || "");
  if (cache.scanRoot?.id === parentId) return cleanPath(cache.scanRoot.remotePath || "");
  return null;
}

function graphParentPath(remotePath = "") {
  const clean = cleanPath(remotePath);
  return `/drive/root:${clean ? `/${clean}` : ""}`;
}

function resolveCachedItemPath(cache, incoming, previous) {
  if (incoming.remotePath !== undefined) return cleanPath(incoming.remotePath);
  const name = incoming.name || previous?.name;
  if (!name) return null;
  const parentId = incoming.parentId;
  const parentPath = cachedParentRemotePath(cache, parentId);
  if (parentPath !== null) return [parentPath, name].filter(Boolean).join("/");
  const sameKnownParent = parentId && previous?.parentId && parentId === previous.parentId;
  const noNewParentEvidence = !parentId;
  if (previous?.remotePath && (sameKnownParent || noNewParentEvidence)) {
    return [remoteParentPath(previous.remotePath), name].filter(Boolean).join("/");
  }
  return null;
}

function updateCachedDescendantPaths(cache, previousPath, nextPath) {
  const oldPath = cleanPath(previousPath || "");
  if (!oldPath) return [];
  const oldLower = cachePathKey(oldPath);
  const changes = [];
  for (const [id, cached] of Object.entries(cache.itemsById || {})) {
    const cachedPath = cleanPath(cached.remotePath || "");
    const cachedLower = cachePathKey(cachedPath);
    if (!cachedPath || !cachedLower.startsWith(`${oldLower}/`)) continue;
    const previous = { ...cached };
    delete cache.pathsByLower[cachedLower];
    if (nextPath === null) {
      delete cached.remotePath;
      delete cached.path;
    } else {
      const suffix = cachedPath.slice(oldPath.length).replace(/^\/+/, "");
      cached.remotePath = [cleanPath(nextPath), suffix].filter(Boolean).join("/");
      cached.path = graphParentPath(remoteParentPath(cached.remotePath));
      cache.pathsByLower[cachePathKey(cached.remotePath)] = id;
    }
    changes.push({ current: cached, previous, removed: [] });
  }
  return changes;
}

function cachePutSimplified(cache, item) {
  if (!item?.id) return { current: null, previous: null, removed: [], descendants: [] };
  if (item.deleted) {
    return { current: null, previous: cache.itemsById?.[item.id] || null, removed: cacheRemoveItemAndDescendants(cache, item), descendants: [] };
  }
  const simplified = Object.hasOwn(item, "remotePath") ? item : simplifyItem(item);
  const previous = cache.itemsById?.[simplified.id] || null;
  const current = mergeDefinedMetadata(previous || {}, simplified);
  const eTagChanged = simplified.eTag !== undefined
    && previous?.eTag !== undefined
    && simplified.eTag !== previous.eTag;
  if (eTagChanged && simplified.cTag === undefined) delete current.cTag;
  if (eTagChanged && simplified.file?.hashes === undefined && current.file) delete current.file.hashes;
  const resolvedPath = resolveCachedItemPath(cache, simplified, previous);
  if (resolvedPath === null) {
    delete current.remotePath;
    delete current.path;
  } else {
    current.remotePath = resolvedPath;
    current.path = simplified.path || graphParentPath(remoteParentPath(resolvedPath));
  }
  const removed = [];
  if (previous?.remotePath && previous.remotePath !== current.remotePath) {
    delete cache.pathsByLower[cachePathKey(previous.remotePath)];
  }
  const pathKey = current.remotePath ? cachePathKey(current.remotePath) : null;
  const existingIdForPath = pathKey ? cache.pathsByLower[pathKey] : null;
  if (existingIdForPath && existingIdForPath !== current.id) {
    const displaced = cache.itemsById[existingIdForPath];
    if (displaced) removed.push(...cacheRemoveItemAndDescendants(cache, displaced));
  }
  const descendants = previous?.folder && previous.remotePath !== current.remotePath
    ? updateCachedDescendantPaths(cache, previous.remotePath, current.remotePath ?? null)
    : [];
  cache.itemsById[current.id] = current;
  if (pathKey) cache.pathsByLower[pathKey] = current.id;
  return { current, previous, removed, descendants };
}

function cacheRemoveItemAndDescendants(cache, item) {
  const old = cache.itemsById?.[item.id] || (item.remotePath ? item : null);
  const remotePath = old?.remotePath || item.remotePath;
  const lowerPath = remotePath ? cachePathKey(remotePath) : null;
  const removed = [];
  const removedIds = new Set([item.id]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const [id, cached] of Object.entries(cache.itemsById || {})) {
      if (removedIds.has(id)) continue;
      const cachedLower = cached.remotePath ? cachePathKey(cached.remotePath) : "";
      if ((lowerPath && (cachedLower === lowerPath || cachedLower.startsWith(`${lowerPath}/`))) || removedIds.has(cached.parentId)) {
        removedIds.add(id);
        changed = true;
      }
    }
  }
  for (const id of removedIds) {
    const cached = cache.itemsById?.[id];
    if (!cached) continue;
    removed.push(cached);
    if (cached.remotePath) delete cache.pathsByLower[cachePathKey(cached.remotePath)];
    delete cache.itemsById[id];
  }
  if (lowerPath) delete cache.pathsByLower[lowerPath];
  return removed;
}

function resolveUnresolvedCachedPaths(cache) {
  const changes = [];
  const maxPasses = Math.max(1, Object.keys(cache.itemsById || {}).length);
  for (let pass = 0; pass < maxPasses; pass += 1) {
    let resolvedThisPass = 0;
    for (const item of Object.values(cache.itemsById || {})) {
      if (item.remotePath !== undefined || !item.name || !item.parentId) continue;
      const parentPath = cachedParentRemotePath(cache, item.parentId);
      if (parentPath === null) continue;
      const previous = { ...item };
      item.remotePath = [parentPath, item.name].filter(Boolean).join("/");
      item.path = graphParentPath(parentPath);
      const pathKey = cachePathKey(item.remotePath);
      const displacedId = cache.pathsByLower[pathKey];
      const removed = [];
      if (displacedId && displacedId !== item.id && cache.itemsById[displacedId]) {
        removed.push(...cacheRemoveItemAndDescendants(cache, cache.itemsById[displacedId]));
      }
      cache.pathsByLower[pathKey] = item.id;
      changes.push({ current: item, previous, removed });
      resolvedThisPass += 1;
    }
    if (!resolvedThisPass) break;
  }
  return changes;
}

async function cacheItems(items = [], metadata = {}) {
  const hasMetadata = metadata.deltaLink !== undefined
    || metadata.deltaNextLink !== undefined
    || metadata.deltaTarget !== undefined
    || metadata.scanRoot !== undefined
    || metadata.pathRoot !== undefined;
  if (!items.length && !hasMetadata) return await loadMetadataCache();
  return await runSerialized("metadata", async () => await withFileLock(metadataCacheLockPath, async () => {
    let cache;
    try {
      cache = await readMetadataCacheFromDisk();
    } catch (error) {
      recordLocalWarning("metadata cache read", error);
      cache = metadataCacheMemory || emptyMetadataCache();
    }
    if (metadata.pathRoot?.id) {
      cache.pathRootsById ||= {};
      cache.pathRootsById[metadata.pathRoot.id] = cleanPath(metadata.pathRoot.remotePath || "");
    }
    if (metadata.scanRoot !== undefined) {
      cache.scanRoot = metadata.scanRoot;
      if (metadata.scanRoot?.id) {
        cache.pathRootsById ||= {};
        cache.pathRootsById[metadata.scanRoot.id] = cleanPath(metadata.scanRoot.remotePath || "");
      }
    }
    const changes = [];
    for (const item of items) {
      const change = cachePutSimplified(cache, item);
      changes.push(change, ...(change.descendants || []));
    }
    changes.push(...resolveUnresolvedCachedPaths(cache));
    await reconcileContentIndexWithMetadata(changes);
    if (metadata.deltaLink !== undefined) {
      cache.deltaLink = metadata.deltaLink || null;
      if (metadata.deltaLink) cache.deltaNextLink = null;
    }
    if (metadata.deltaNextLink !== undefined) cache.deltaNextLink = metadata.deltaNextLink || null;
    if (metadata.deltaTarget !== undefined) cache.deltaTarget = metadata.deltaTarget || null;
    return await saveMetadataCache(cache);
  }));
}

async function clearMetadataCache() {
  return await runSerialized("metadata", async () => await withFileLock(metadataCacheLockPath, async () => {
    metadataCacheMemory = emptyMetadataCache();
    metadataCacheLoadPromise = null;
    await ensurePrivateDirectory(cacheRoot);
    await writePrivateFileAtomic(cachePath, JSON.stringify(metadataCacheMemory));
    metadataCacheFileVersion = await localFileVersion(cachePath);
    return metadataCacheMemory;
  }));
}

function emptyContentIndex() {
  return {
    version: 2,
    createdAt: new Date().toISOString(),
    updatedAt: null,
    itemCount: 0,
    entriesById: {}
  };
}

async function loadContentIndex() {
  if (!contentIndexLoadPromise) {
    contentIndexLoadPromise = (async () => {
      try {
        const diskVersion = await localFileVersion(contentIndexPath);
        if (contentIndexMemory && diskVersion === contentIndexFileVersion) return contentIndexMemory;
        await withFileLock(contentIndexLockPath, async () => {
          const lockedVersion = await localFileVersion(contentIndexPath);
          if (contentIndexMemory && lockedVersion === contentIndexFileVersion) return;
          contentIndexMemory = await readContentIndexFromDisk();
          contentIndexFileVersion = await localFileVersion(contentIndexPath);
        });
      } catch (error) {
        recordLocalWarning("content index read", error);
        contentIndexMemory = emptyContentIndex();
        contentIndexFileVersion = null;
      } finally {
        contentIndexLoadPromise = null;
      }
      return contentIndexMemory;
    })();
  }
  return await contentIndexLoadPromise;
}

async function readContentIndexFromDisk() {
  await ensurePrivateDirectory(cacheRoot);
  try {
    const parsed = JSON.parse(await readFile(contentIndexPath, "utf8"));
    await hardenPrivateFile(contentIndexPath);
    return {
      ...emptyContentIndex(),
      ...parsed,
      entriesById: parsed.entriesById || {}
    };
  } catch (error) {
    if (error?.code === "ENOENT") return emptyContentIndex();
    throw error;
  }
}

async function writeContentIndex(index) {
  index.updatedAt = new Date().toISOString();
  index.itemCount = Object.keys(index.entriesById || {}).length;
  await ensurePrivateDirectory(cacheRoot);
  await writePrivateFileAtomic(contentIndexPath, JSON.stringify(index));
  contentIndexMemory = index;
  contentIndexFileVersion = await localFileVersion(contentIndexPath);
  return index;
}

async function saveContentIndex(index) {
  return await runSerialized("content", async () => await withFileLock(contentIndexLockPath, async () => {
    let latest;
    try {
      latest = await readContentIndexFromDisk();
    } catch (error) {
      recordLocalWarning("content index read", error);
      latest = contentIndexMemory || emptyContentIndex();
    }
    const merged = {
      ...latest,
      entriesById: {
        ...(latest.entriesById || {}),
        ...(index.entriesById || {})
      }
    };
    return await writeContentIndex(merged);
  }));
}

async function clearContentIndex() {
  return await runSerialized("content", async () => await withFileLock(contentIndexLockPath, async () => {
    contentIndexMemory = emptyContentIndex();
    contentIndexLoadPromise = null;
    await ensurePrivateDirectory(cacheRoot);
    await writePrivateFileAtomic(contentIndexPath, JSON.stringify(contentIndexMemory));
    contentIndexFileVersion = await localFileVersion(contentIndexPath);
    return contentIndexMemory;
  }));
}

function indexedItemMetadataChanged(left = {}, right = {}) {
  return left.id !== right.id
    || left.name !== right.name
    || left.remotePath !== right.remotePath
    || left.path !== right.path
    || left.webUrl !== right.webUrl
    || left.size !== right.size
    || left.createdDateTime !== right.createdDateTime
    || left.lastModifiedDateTime !== right.lastModifiedDateTime
    || left.eTag !== right.eTag
    || left.cTag !== right.cTag
    || left.parentId !== right.parentId
    || left.driveId !== right.driveId
    || Boolean(left.deleted) !== Boolean(right.deleted);
}

async function reconcileContentIndexWithMetadata(changes = []) {
  if (!changes.length) return;
  await runSerialized("content", async () => await withFileLock(contentIndexLockPath, async () => {
    let index;
    try {
      index = await readContentIndexFromDisk();
    } catch (error) {
      recordLocalWarning("content index read", error);
      index = contentIndexMemory || emptyContentIndex();
    }
    let changed = false;
    for (const change of changes) {
      for (const removed of change?.removed || []) {
        if (removed?.id && Object.hasOwn(index.entriesById, removed.id)) {
          delete index.entriesById[removed.id];
          changed = true;
        }
      }
      const current = change?.current;
      if (!current?.id) continue;
      const entry = index.entriesById[current.id];
      if (!entry) continue;
      if (!contentIndexEntryFresh(entry, current)) {
        delete index.entriesById[current.id];
        changed = true;
        continue;
      }
      if (indexedItemMetadataChanged(entry.item, current)) {
        entry.item = current;
        changed = true;
      }
    }
    contentIndexMemory = index;
    if (changed) await writeContentIndex(index);
  }));
}

function contentIndexEntryFresh(entry, item) {
  if (!entry) return false;
  if (entry.cTag && item.cTag) return entry.cTag === item.cTag;
  return entry.eTag === item.eTag
    && entry.lastModifiedDateTime === item.lastModifiedDateTime
    && entry.size === item.size;
}

function officeExportIndexable(item = {}) {
  const extension = extname(item.name || "").toLowerCase();
  const mimeType = (item.file?.mimeType || "").toLowerCase();
  return Object.values(officeKinds).some((kind) => kind.extensions.has(extension) || kind.mimeTypes.has(mimeType));
}

function contentIndexableReason(item = {}, args = {}) {
  if (!item?.file || item.folder) return { ok: false, reason: "not-file" };
  const settings = pluginSettings();
  const maxBytes = clampInteger(args.maxBytesPerFile, settings.maxIndexedFileSize, 1024, textFileLimit);
  const isOffice = officeExportIndexable(item);
  if (args.officeOnly === true && !isOffice) return { ok: false, reason: "not-office" };
  if (isOffice && args.includeOfficeStructured !== false) {
    const officeIndexLimit = clampInteger(args.maxOfficePackageBytes, 50 * 1024 * 1024, 1024, maxOfficePackageBytes);
    if (item.size && item.size > officeIndexLimit) return { ok: false, reason: "office-package-too-large" };
    return { ok: true, source: "office-openxml", kind: officePackageKindFromName(item.name) };
  }
  if (item.size && item.size > maxBytes) return { ok: false, reason: "too-large" };
  const allowedExtensions = normalizeExtensions(args.extensions?.length ? args.extensions : settings.supportedIndexedFileTypes);
  const extension = extname(item.name || "").toLowerCase();
  if (allowedExtensions.has(extension) || isLikelyTextItem(item)) return { ok: true, source: "text-read" };
  if ((args.includeOfficeExport === true || settings.includeOfficeTextExport) && officeExportIndexable(item)) {
    return { ok: true, source: "graph-text-export" };
  }
  return { ok: false, reason: "unsupported-type" };
}

function officeIndexSegments(document, maxSegments = 50_000) {
  const segments = [];
  const add = (text, anchor) => {
    const value = String(text ?? "").trim();
    if (!value || segments.length >= maxSegments) return;
    segments.push({ text: value, anchor });
  };
  if (document.kind === "word") {
    for (const paragraph of document.paragraphs || []) add(paragraph.text, { type: "paragraph", part: paragraph.part, index: paragraph.index, style: paragraph.style || null });
    for (const table of document.tables || []) {
      for (const [rowIndex, row] of (table.rows || []).entries()) {
        for (const [columnIndex, value] of row.entries()) add(value, { type: "tableCell", part: table.part, tableIndex: table.index, rowIndex, columnIndex });
      }
    }
    for (const control of document.contentControls || []) add(control.text, { type: "contentControl", part: control.part, index: control.index, id: control.id, tag: control.tag });
    for (const comment of document.comments || []) add(comment.text, { type: "comment", id: comment.id, author: comment.author });
  } else if (document.kind === "excel") {
    for (const sheet of document.sheets || []) {
      for (const cell of sheet.cells || []) {
        add(cell.value, { type: "cell", sheet: sheet.name, address: cell.address, formula: cell.formula || null });
        if (cell.formula) add(cell.formula, { type: "formula", sheet: sheet.name, address: cell.address });
      }
      for (const table of sheet.tables || []) add(`${table.name || table.displayName || "table"} ${table.ref || ""}`, { type: "table", sheet: sheet.name, name: table.name, ref: table.ref });
    }
    for (const name of document.definedNames || []) add(`${name.name || ""} ${name.value || ""}`, { type: "definedName", name: name.name });
  } else if (document.kind === "powerpoint") {
    const addShape = (shape, slideIndex, parentShapeId = null) => {
      add(shape.text, { type: "shape", slideIndex, shapeId: shape.id, name: shape.name, parentShapeId });
      for (const [rowIndex, row] of (shape.table?.rows || []).entries()) {
        for (const [columnIndex, value] of row.entries()) add(value, { type: "tableCell", slideIndex, shapeId: shape.id, rowIndex, columnIndex });
      }
      for (const child of shape.children || []) addShape(child, slideIndex, shape.id);
    };
    for (const slide of document.slides || []) {
      for (const shape of slide.shapes || []) addShape(shape, slide.index);
      add(slide.notes, { type: "notes", slideIndex: slide.index });
    }
  }
  return { segments, truncated: segments.length >= maxSegments };
}

async function extractIndexText(item, args = {}) {
  const settings = pluginSettings();
  const maxBytes = clampInteger(args.maxBytesPerFile, settings.maxIndexedFileSize, 1024, textFileLimit);
  const indexable = contentIndexableReason(item, args);
  if (!indexable.ok) throw new Error(indexable.reason);
  if (indexable.source === "office-openxml") {
    const document = await inspectRemoteOfficePackage({
      itemId: item.id,
      maxParagraphs: clampInteger(args.maxOfficeParagraphs, 10_000, 1, 10_000),
      maxCells: clampInteger(args.maxOfficeCells, 50_000, 1, 50_000),
      maxSlides: clampInteger(args.maxOfficeSlides, 5_000, 1, 5_000),
      includeCells: true,
      strictRelationships: true
    }, indexable.kind, "inspect");
    const structured = officeIndexSegments(document, clampInteger(args.maxOfficeSegments, 50_000, 1, 100_000));
    const text = structured.segments.map((entry) => entry.text).join("\n");
    return {
      text,
      segments: structured.segments,
      structuredKind: document.kind,
      bytesRead: Number(item.size || 0),
      truncated: structured.truncated || Boolean(document.truncated),
      source: "office-openxml"
    };
  }
  const target = { itemId: item.id };
  const sourcePath = indexable.source === "graph-text-export"
    ? `${contentPath(target)}?${new URLSearchParams({ format: "text" }).toString()}`
    : contentPath(target);
  const limited = await graphLimitedBuffer(sourcePath, maxBytes);
  if (indexable.source === "text-read") assertNoBinaryNulls(limited.buffer);
  return {
    text: limited.buffer.toString("utf8").replace(/\uFFFD$/u, ""),
    bytesRead: limited.bytesRead,
    truncated: limited.truncated,
    source: indexable.source
  };
}

function contentIndexTextFields(text = "") {
  return {
    normalizedText: normalizeFindText(text),
    tokens: findTokens(text)
  };
}

function contentSearchContext(query) {
  return {
    query,
    normalizedQuery: normalizeFindText(query),
    tokens: findImportantTokens(query)
  };
}

function compareContentMatches(left, right) {
  if (right.score !== left.score) return right.score - left.score;
  return String(left.item?.name || "").localeCompare(String(right.item?.name || ""));
}

function insertTopContentMatch(matches, candidate, maxResults) {
  matches.push(candidate);
  matches.sort(compareContentMatches);
  if (matches.length > maxResults) matches.pop();
}

function contentMatchForQuery(entry, queryOrContext) {
  const context = typeof queryOrContext === "object" && queryOrContext
    ? queryOrContext
    : contentSearchContext(queryOrContext);
  const text = String(entry.text || "");
  const normalizedText = entry.normalizedText || normalizeFindText(text);
  const normalizedQuery = context.normalizedQuery;
  const tokens = context.tokens || [];
  const textTokenSet = new Set(Array.isArray(entry.tokens) ? entry.tokens : findTokens(text));
  const tokenMatches = tokens.filter((token) => textTokenSet.has(token));
  if (!normalizedQuery && !tokenMatches.length) return null;
  const exactPhrase = normalizedQuery
    ? (/\s|[-.]/.test(normalizedQuery) ? normalizedText.includes(normalizedQuery) : textTokenSet.has(normalizedQuery))
    : false;
  const exactIndex = exactPhrase ? normalizedText.indexOf(normalizedQuery) : -1;
  if (exactIndex < 0 && tokenMatches.length === 0) return null;

  const lowerText = text.toLowerCase();
  const rawNeedle = exactIndex >= 0 ? normalizedQuery : tokenMatches[0];
  const rawIndex = rawNeedle ? lowerText.indexOf(rawNeedle.toLowerCase()) : -1;
  const start = rawIndex >= 0 ? Math.max(0, rawIndex - 80) : 0;
  const end = rawIndex >= 0 ? Math.min(text.length, rawIndex + rawNeedle.length + 120) : Math.min(text.length, 200);
  const snippet = text.slice(start, end).replace(/\s+/g, " ").trim();
  const coverage = tokens.length ? tokenMatches.length / tokens.length : 0;
  const segmentMatches = [];
  for (const segment of entry.segments || []) {
    const segmentText = String(segment.text || "");
    const normalizedSegment = normalizeFindText(segmentText);
    const segmentTokens = new Set(findTokens(segmentText));
    const phrase = normalizedQuery && normalizedSegment.includes(normalizedQuery);
    const matched = tokens.filter((token) => segmentTokens.has(token));
    if (!phrase && !matched.length) continue;
    segmentMatches.push({
      anchor: segment.anchor,
      snippet: segmentText.slice(0, 500),
      score: (phrase ? 70 : 35) + Math.round((tokens.length ? matched.length / tokens.length : 0) * 35)
    });
  }
  segmentMatches.sort((a, b) => b.score - a.score);
  return {
    score: (exactIndex >= 0 ? 70 : 35) + Math.round(coverage * 35),
    matchedTokens: tokenMatches.length,
    snippet: segmentMatches[0]?.snippet || snippet,
    anchor: segmentMatches[0]?.anchor || null,
    segmentMatches: segmentMatches.slice(0, 10),
    reasons: [
      exactIndex >= 0 ? "indexed content phrase" : "indexed content tokens",
      ...(tokenMatches.length ? [`content token matches: ${tokenMatches.slice(0, 5).join(", ")}`] : [])
    ]
  };
}

async function contentSearch(args = {}) {
  const startedAt = Date.now();
  const query = String(args.query || "").trim();
  if (!query) throw new Error("query is required.");
  const maxResults = clampInteger(args.maxResults, 20, 1, 100);
  const index = await loadContentIndex();
  const context = contentSearchContext(query);
  const matches = [];
  let matched = 0;
  for (const entry of Object.values(index.entriesById || {})) {
    if (args.officeOnly === true && entry.source !== "office-openxml") continue;
    const match = contentMatchForQuery(entry, context);
    if (!match) continue;
    matched += 1;
    insertTopContentMatch(matches, {
      item: entry.item,
      score: match.score,
      matchedTokens: match.matchedTokens,
      snippet: match.snippet,
      anchor: match.anchor,
      segmentMatches: match.segmentMatches,
      reasons: match.reasons,
      source: entry.source,
      indexedAt: entry.indexedAt,
      truncated: entry.truncated
    }, maxResults);
  }
  return {
    query,
    indexPath: contentIndexPath,
    itemCount: index.itemCount || 0,
    matched,
    returned: matches.length,
    durationMs: elapsedMs(startedAt),
    items: matches.map((match) => ({
      ...formatSimplifiedItem(match.item, args.format),
      score: match.score,
      matchedTokens: match.matchedTokens,
      reasons: match.reasons,
      snippet: match.snippet,
      anchor: match.anchor,
      segmentMatches: match.segmentMatches,
      source: match.source,
      indexedAt: match.indexedAt,
      truncated: match.truncated
    }))
  };
}

async function officeIndexRefresh(args = {}) {
  return await contentIndexRefresh({
    ...args,
    refreshMetadata: args.refreshMetadata !== false,
    metadataMode: args.metadataMode || "auto",
    includeOfficeStructured: true,
    officeOnly: true
  });
}

async function officeContentSearch(args = {}) {
  return await contentSearch({ ...args, officeOnly: true });
}

async function contentIndexRefresh(args = {}) {
  const startedAt = Date.now();
  const settings = pluginSettings();
  if (!settings.contentIndexEnabled) throw new Error("Content indexing is disabled by configuration.");
  if (args.refreshMetadata === true) {
    await cacheRefresh({
      ...args,
      mode: args.metadataMode || "auto",
      maxItems: clampInteger(args.scanMaxItems, 10000, 1, 50000),
      maxFolders: clampInteger(args.scanMaxFolders, 2000, 1, 10000),
      maxDepth: clampInteger(args.scanMaxDepth, settings.maxScanDepth, 0, 50)
    });
  }

  const maxFiles = clampInteger(args.maxFiles, 100, 1, 1000);
  let sourceItems = [];
  if (args.itemId || args.path || args.preset) {
    sourceItems = [simplifyItem(await getRawInfo(args))];
  } else {
    sourceItems = await cachedItems();
  }

  const index = await loadContentIndex();
  const eligibleItems = sourceItems
    .filter((item) => item?.id && !item.deleted)
    .filter((item) => contentIndexableReason(item, args).ok);
  const candidates = eligibleItems
    .map((item) => ({
      item,
      fresh: Boolean(args.force !== true && contentIndexEntryFresh(index.entriesById[item.id], item))
    }))
    .sort((left, right) => Number(left.fresh) - Number(right.fresh))
    .slice(0, maxFiles)
    .map((entry) => entry.item);
  const results = {
    indexPath: contentIndexPath,
    considered: sourceItems.length,
    eligible: eligibleItems.length,
    selected: candidates.length,
    indexed: 0,
    reused: 0,
    skipped: sourceItems.length - candidates.length,
    graphContentReadsAttempted: 0,
    failed: 0,
    failures: []
  };

  await mapWithConcurrency(candidates, clampInteger(args.concurrencyLimit, settings.concurrencyLimit, 1, 8), async (item) => {
    const existing = index.entriesById[item.id];
    if (args.force !== true && contentIndexEntryFresh(existing, item)) {
      index.entriesById[item.id] = { ...existing, item };
      results.reused += 1;
      return;
    }
    try {
      results.graphContentReadsAttempted += 1;
      const extracted = await extractIndexText(item, args);
      const textFields = contentIndexTextFields(extracted.text);
      index.entriesById[item.id] = {
        item,
        text: extracted.text,
        normalizedText: textFields.normalizedText,
        tokens: textFields.tokens,
        indexedAt: new Date().toISOString(),
        source: extracted.source,
        bytesRead: extracted.bytesRead,
        textBytes: Buffer.byteLength(extracted.text, "utf8"),
        truncated: extracted.truncated,
        segments: extracted.segments || [],
        structuredKind: extracted.structuredKind || null,
        eTag: item.eTag,
        cTag: item.cTag,
        lastModifiedDateTime: item.lastModifiedDateTime,
        size: item.size
      };
      results.indexed += 1;
    } catch (error) {
      results.failed += 1;
      if (results.failures.length < 10) {
        results.failures.push({ id: item.id, name: item.name, error: safeToolErrorMessage(error) });
      }
    }
  });

  const updatedEntries = Object.fromEntries(
    candidates
      .filter((item) => index.entriesById[item.id])
      .map((item) => [item.id, index.entriesById[item.id]])
  );
  const persistedIndex = await saveContentIndex({ entriesById: updatedEntries });
  return {
    ...results,
    itemCount: persistedIndex.itemCount,
    durationMs: elapsedMs(startedAt),
    settings: {
      maxBytesPerFile: clampInteger(args.maxBytesPerFile, settings.maxIndexedFileSize, 1024, textFileLimit),
      supportedIndexedFileTypes: args.extensions?.length ? [...normalizeExtensions(args.extensions)] : settings.supportedIndexedFileTypes,
      includeOfficeExport: args.includeOfficeExport === true || settings.includeOfficeTextExport,
      includeOfficeStructured: args.includeOfficeStructured !== false,
      concurrencyLimit: clampInteger(args.concurrencyLimit, settings.concurrencyLimit, 1, 8)
    },
    note: "Content indexing reads file bodies only during this explicit refresh. Normal find/search calls reuse the local index and do not fetch content."
  };
}

async function cacheMovedOrRenamedItem(previous, current) {
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([current]));
}

function unresolvedPathItems(cache) {
  return Object.values(cache.itemsById || {}).filter((item) => item?.id && !item.deleted && item.remotePath === undefined);
}

async function syncStatus(args = {}) {
  const cache = await loadMetadataCache();
  const contentIndex = await loadContentIndex();
  const settings = pluginSettings();
  const items = Object.values(cache.itemsById || {});
  const contentEntries = Object.values(contentIndex.entriesById || {});
  const unresolvedItems = unresolvedPathItems(cache);
  const cacheAgeSeconds = cache.updatedAt ? Math.max(0, Math.round((Date.now() - Date.parse(cache.updatedAt)) / 1000)) : null;
  return {
    cachePath,
    contentIndexPath,
    downloadRoot,
    updateRoot,
    backupRoot,
    storageRoot,
    itemCount: items.length,
    updatedAt: cache.updatedAt,
    cacheAgeSeconds,
    cacheFresh: cacheAgeSeconds === null ? false : settings.cacheTtlSeconds === 0 || cacheAgeSeconds <= settings.cacheTtlSeconds,
    deltaLinkAvailable: Boolean(cache.deltaLink),
    deltaNextLinkAvailable: Boolean(cache.deltaNextLink),
    deltaTarget: cache.deltaTarget || null,
    scanRoot: cache.scanRoot,
    unresolvedPathCount: unresolvedItems.length,
    unresolvedPathSamples: args.includeSamples
      ? unresolvedItems.slice(0, 10).map((item) => ({ id: item.id, name: item.name, parentId: item.parentId }))
      : undefined,
    contentIndex: {
      itemCount: contentEntries.length,
      structuredOfficeCount: contentEntries.filter((entry) => entry.source === "office-openxml").length,
      segmentCount: contentEntries.reduce((sum, entry) => sum + (entry.segments?.length || 0), 0),
      updatedAt: contentIndex.updatedAt,
      enabled: settings.contentIndexEnabled
    },
    settings,
    samples: args.includeSamples ? items.slice(0, 10).map((item) => formatSimplifiedItem(item, "compact")) : undefined
  };
}

async function cachedItems() {
  const cache = await loadMetadataCache();
  return Object.values(cache.itemsById || {});
}

async function cachedItemByPath(path) {
  const cache = await loadMetadataCache();
  const id = cache.pathsByLower?.[cachePathKey(path)];
  return id ? cache.itemsById[id] : null;
}

async function cachedItemById(id) {
  const cache = await loadMetadataCache();
  return id ? cache.itemsById[id] || null : null;
}

async function mapWithConcurrency(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    for (;;) {
      const index = next;
      next += 1;
      if (index >= items.length) return;
      results[index] = await fn(items[index], index);
    }
  });
  await Promise.all(workers);
  return results;
}

async function doctor(args = {}) {
  const checks = [];
  const addCheck = (name, status, details = {}) => {
    checks.push({ name, status, details });
  };
  const runCheck = async (name, fn) => {
    try {
      addCheck(name, "pass", await fn());
    } catch (error) {
      addCheck(name, "fail", { error: safeToolErrorMessage(error) });
    }
  };

  const cfgStatus = publicConfig();
  addCheck("config", cfgStatus.clientIdConfigured ? "pass" : "fail", {
    clientIdConfigured: cfgStatus.clientIdConfigured,
    tenant: cfgStatus.tenant,
    scopes: cfgStatus.scopes,
    keychainTokenConfigured: cfgStatus.keychainTokenConfigured,
    configPath: cfgStatus.configPath
  });

  await runCheck("access token", async () => {
    await getAccessToken();
    return { accessTokenAvailable: true };
  });
  await runCheck("profile", async () => {
    const profile = await graph("/me");
    return {
      displayName: profile.displayName,
      userPrincipalName: profile.userPrincipalName,
      mail: profile.mail
    };
  });
  await runCheck("drive", async () => {
    const drive = await graph("/me/drive");
    return {
      id: drive.id,
      name: drive.name,
      driveType: drive.driveType,
      quotaState: drive.quota?.state
    };
  });
  addCheck("presets", "pass", { pathPresets: cfgStatus.pathPresets });

  if (args.checkRootList !== false) {
    await runCheck("root list", async () => {
      const result = await list({ limit: Math.min(args.rootListLimit ?? 5, 20) });
      return {
        returned: result.items.length,
        nextLink: Boolean(result.nextLink),
        sample: result.items.slice(0, 3).map((item) => ({ name: item.name, type: item.type }))
      };
    });
  }

  const failCount = checks.filter((check) => check.status === "fail").length;
  const warnCount = checks.filter((check) => check.status === "warn").length;
  return {
    ok: failCount === 0,
    status: failCount ? "fail" : warnCount ? "warn" : "pass",
    summary: { total: checks.length, pass: checks.length - failCount - warnCount, warn: warnCount, fail: failCount },
    checks,
    note: failCount
      ? "At least one OneDrive health check failed. Fix the failed check before relying on file tools."
      : "OneDrive plugin health checks passed."
  };
}

function normalizeToken(raw) {
  const now = Date.now();
  return {
    ...raw,
    expires_at: raw.expires_at || (raw.expires_in ? now + raw.expires_in * 1000 : null),
    obtained_at: raw.obtained_at || now
  };
}

async function postForm(url, params) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
    signal: timeoutSignal()
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = body.error_description || body.error || `${response.status} ${response.statusText}`;
    const error = new Error(message);
    error.body = body;
    throw error;
  }
  return body;
}

async function startDeviceLogin(args = {}) {
  const generation = authGeneration;
  const deviceGeneration = ++deviceLoginGeneration;
  const cfg = config(args);
  requireClientId(cfg);
  const { body: result, tenant } = await postFormWithConsumerFallback("device", cfg, {
    client_id: cfg.clientId,
    scope: cfg.scopes
  });
  if (generation !== authGeneration || deviceGeneration !== deviceLoginGeneration) {
    throw new Error("OneDrive authentication state changed while device login was starting. Start login again.");
  }
  pendingDevice = { ...result, tenant, scopes: cfg.scopes, startedAt: Date.now(), deviceGeneration };
  return {
    userCode: result.user_code,
    verificationUri: result.verification_uri,
    verificationUriComplete: result.verification_uri_complete,
    expiresIn: result.expires_in,
    interval: result.interval,
    message: result.message,
    authTenant: tenant,
    deviceCodeStoredInMemory: true
  };
}

async function pollDeviceLogin(args = {}) {
  const generation = authGeneration;
  const pending = pendingDevice;
  const deviceCode = args.deviceCode || pending?.device_code;
  if (!deviceCode) throw new Error("No pending device code. Run onedrive_auth_device_start first.");
  if (pending && args.deviceCode && args.deviceCode !== pending.device_code) {
    throw new Error("The supplied device code does not match the active OneDrive device-login session. Start login again.");
  }
  const deviceGeneration = pending?.deviceGeneration ?? deviceLoginGeneration;
  const cfg = config({ tenant: pending?.tenant, scopes: pending?.scopes });
  requireClientId(cfg);
  const tokenParams = {
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    client_id: cfg.clientId,
    device_code: deviceCode
  };
  const tokenResponse = await postFormWithConsumerFallback("token", cfg, tokenParams).catch((error) => {
    if (error.body?.error === "authorization_pending") {
      return { authorizationPending: true, message: "Authorization is still pending. Try again after the user completes browser sign-in." };
    }
    if (error.body?.error === "slow_down") {
      return { authorizationPending: true, slowDown: true, message: "Microsoft asked polling to slow down. Try again in a few more seconds." };
    }
    throw error;
  });
  const result = tokenResponse.body || tokenResponse;
  if (generation !== authGeneration
    || deviceGeneration !== deviceLoginGeneration
    || (pending && pendingDevice?.deviceGeneration !== deviceGeneration)) {
    throw new Error("OneDrive authentication state changed while device login was pending. Start login again.");
  }
  if (result.authorizationPending) return result;
  tokenCache = normalizeToken(result);
  setKeychainToken(tokenCache, cfg);
  pendingDevice = null;
  return {
    authenticated: true,
    authTenant: tokenResponse.tenant || cfg.tenant,
    tokenType: tokenCache.token_type,
    expiresAt: tokenCache.expires_at ? new Date(tokenCache.expires_at).toISOString() : null,
    refreshTokenStoredInKeychain: Boolean(tokenCache.refresh_token)
  };
}

async function refreshAccessToken(refreshToken, cfg = config(), generation = authGeneration) {
  requireClientId(cfg);
  const { body: result, tenant } = await postFormWithConsumerFallback("token", cfg, {
    grant_type: "refresh_token",
    client_id: cfg.clientId,
    refresh_token: refreshToken,
    scope: cfg.scopes
  });
  if (generation !== authGeneration) {
    throw new Error("OneDrive authentication state changed while the access token was refreshing. Try again if access is still intended.");
  }
  tokenCache = normalizeToken({ ...result, auth_tenant: tenant, refresh_token: result.refresh_token || refreshToken });
  setKeychainToken(tokenCache, cfg);
  return tokenCache;
}

async function getAccessToken() {
  if (process.env.ONEDRIVE_TEST_ACCESS_TOKEN) return process.env.ONEDRIVE_TEST_ACCESS_TOKEN;
  const cfg = config();
  requireClientId(cfg);
  const current = tokenCache || getKeychainToken(cfg);
  if (!current?.refresh_token && !current?.access_token) {
    throw new Error("OneDrive is not authenticated. Run onedrive_auth_device_start, complete browser login, then run onedrive_auth_device_poll.");
  }
  const expiresAt = current.expires_at || 0;
  if (current.access_token && expiresAt - Date.now() > 60_000) {
    tokenCache = current;
    return current.access_token;
  }
  if (!current.refresh_token) throw new Error("Stored token has no refresh token. Run device-code login again.");
  if (!tokenRefreshPromise) {
    const generation = authGeneration;
    tokenRefreshPromise = refreshAccessToken(current.refresh_token, cfg, generation);
  }
  const refresh = tokenRefreshPromise;
  try {
    const refreshed = await refresh;
    return refreshed.access_token;
  } finally {
    if (tokenRefreshPromise === refresh) tokenRefreshPromise = null;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function elapsedMs(startedAt) {
  return Math.max(0, Date.now() - startedAt);
}

function fetchTimeoutMs() {
  const value = Number(process.env.ONEDRIVE_FETCH_TIMEOUT_MS || 60_000);
  return Number.isFinite(value) && value > 0 ? value : 60_000;
}

function timeoutSignal(timeoutMs = fetchTimeoutMs()) {
  if (AbortSignal.timeout) return AbortSignal.timeout(timeoutMs);
  const controller = new AbortController();
  setTimeout(() => controller.abort(), timeoutMs).unref?.();
  return controller.signal;
}

function retryDelayMs(response, attempt) {
  const retryAfter = response.headers.get("retry-after");
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const dateMs = Date.parse(retryAfter);
    if (Number.isFinite(dateMs)) return Math.max(0, dateMs - Date.now());
  }
  return Math.min(1000 * 2 ** attempt, 8000);
}

function shouldRetryResponse(response) {
  return response.status === 429 || response.status === 503 || response.status === 504 || response.status === 502 || response.status === 500;
}

function isReplayableBody(body) {
  return body === undefined
    || body === null
    || typeof body === "string"
    || Buffer.isBuffer(body)
    || body instanceof ArrayBuffer
    || body instanceof URLSearchParams;
}

function retryCountForRequest(fetchOptions, explicitMaxRetries) {
  if (explicitMaxRetries === 0) return 0;
  const method = String(fetchOptions.method || "GET").toUpperCase();
  if (method === "GET" || method === "HEAD") return explicitMaxRetries ?? 4;
  if (explicitMaxRetries !== undefined && isReplayableBody(fetchOptions.body)) return explicitMaxRetries;
  return 0;
}

function shouldDefaultJsonContentType(body) {
  return body
    && typeof body === "string"
    && body.trim().startsWith("{");
}

function contentLength(response) {
  const value = Number(response.headers.get("content-length"));
  return Number.isFinite(value) && value >= 0 ? value : null;
}

async function parseResponseBody(response) {
  if (response.status === 204) return null;
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) return await response.json().catch(() => ({}));
  const buffer = await response.arrayBuffer();
  if (contentType.startsWith("text/") || (!response.ok && buffer.byteLength <= 4096)) {
    return new TextDecoder().decode(buffer);
  }
  return buffer;
}

async function fetchWithRetry(url, options = {}, retryOptions = {}) {
  const maxRetries = retryOptions.maxRetries ?? 3;
  const timeoutMs = retryOptions.timeoutMs ?? fetchTimeoutMs();
  for (let attempt = 0; ; attempt += 1) {
    let response;
    try {
      response = await fetch(url, { ...options, signal: options.signal || timeoutSignal(timeoutMs) });
    } catch (error) {
      if (attempt >= maxRetries) {
        if (error.name === "AbortError" || error.name === "TimeoutError") {
          throw new Error(`Request timed out after ${timeoutMs}ms: ${url}`);
        }
        const transportError = new Error(`Microsoft Graph transport error after ${attempt + 1} attempts: ${error.message || String(error)}`);
        transportError.cause = error;
        throw transportError;
      }
      await sleep(Math.min(1000 * 2 ** attempt, 8000));
      continue;
    }
    if (!shouldRetryResponse(response) || attempt >= maxRetries) return response;
    await parseResponseBody(response).catch(() => null);
    await sleep(retryDelayMs(response, attempt));
  }
}

function graphRequestId(headers, body) {
  return headers?.get?.("request-id")
    || headers?.get?.("x-ms-request-id")
    || headers?.get?.("client-request-id")
    || body?.error?.innerError?.["request-id"]
    || body?.error?.innerError?.requestId
    || null;
}

function currentGraphRequestId(options = {}) {
  const store = toolCallContext.getStore();
  return options.mutation === true
    ? store?.lastMutationGraphRequestId || null
    : store?.lastGraphRequestId || null;
}

function rememberGraphRequestId(requestId, options = {}) {
  const store = toolCallContext.getStore();
  if (!store) return;
  store.lastGraphRequestId = requestId || null;
  if (options.mutation === true) store.lastMutationGraphRequestId = requestId || null;
}

function microsoftGraphError(body, response) {
  const code = body?.error?.code ? `${body.error.code}: ` : "";
  const textBody = typeof body === "string" ? body.trim() : "";
  const message = body?.error?.message || textBody || `${response.status} ${response.statusText}`;
  const requestId = graphRequestId(response.headers, body);
  const suffix = requestId ? ` (request-id: ${requestId})` : "";
  const error = new Error(`Microsoft Graph error: ${code}${message}${suffix}`);
  error.graphRequestId = requestId;
  error.graphStatus = response.status;
  return error;
}

function graphBaseUrl() {
  return process.env.ONEDRIVE_GRAPH_BASE_URL || "https://graph.microsoft.com/v1.0";
}

function graphUrl(path, options = {}) {
  const base = new URL(graphBaseUrl());
  if (!String(path).startsWith("http")) return `${base.toString().replace(/\/+$/, "")}${path}`;
  const target = new URL(path);
  if (options.skipAuth) return target.toString();
  const basePath = base.pathname.replace(/\/+$/, "");
  const targetAllowed = target.origin === base.origin
    && (target.pathname === basePath || target.pathname.startsWith(`${basePath}/`));
  if (!targetAllowed) {
    throw new Error(`Refusing to send an authenticated Microsoft Graph request to an untrusted URL: ${target.origin}${target.pathname}`);
  }
  return target.toString();
}

function assertTrustedCopyMonitorUrl(monitorUrl) {
  const base = new URL(graphBaseUrl());
  const target = new URL(monitorUrl, base);
  if (target.origin === base.origin) return target.toString();
  const host = target.hostname.toLowerCase();
  const trustedMicrosoftHost = host === "my.microsoftpersonalcontent.com"
    || host.endsWith(".microsoftpersonalcontent.com")
    || host.endsWith(".sharepoint.com");
  if (target.protocol !== "https:" || !trustedMicrosoftHost) {
    throw new Error(`Refusing to poll untrusted copy monitor URL: ${target.origin}${target.pathname}`);
  }
  return target.toString();
}

function assertTrustedUploadSessionUrl(uploadUrl) {
  const base = new URL(graphBaseUrl());
  const target = new URL(uploadUrl, base);
  if (target.origin === base.origin) return target.toString();
  const host = target.hostname.toLowerCase();
  const trustedMicrosoftHost = host.endsWith(".up.1drv.com")
    || host.endsWith(".up.1drv.ms")
    || host.endsWith(".sharepoint.com")
    || host.endsWith(".microsoftpersonalcontent.com");
  if (target.protocol !== "https:" || !trustedMicrosoftHost) {
    throw new Error(`Refusing to upload file contents to an untrusted upload session URL: ${target.origin}${target.pathname}`);
  }
  return target.toString();
}

async function graph(path, options = {}) {
  const { returnResponse = false, maxRetries, skipAuth = false, isMutation, ...fetchOptions } = options;
  const method = String(fetchOptions.method || "GET").toUpperCase();
  const mutationRequest = isMutation ?? !["GET", "HEAD", "OPTIONS"].includes(method);
  const accessToken = skipAuth ? null : await getAccessToken();
  const url = graphUrl(path, { skipAuth });
  const headers = {
    ...(accessToken ? { Authorization: `Bearer ${accessToken}` } : {}),
    ...(shouldDefaultJsonContentType(fetchOptions.body) ? { "Content-Type": "application/json" } : {}),
    ...(fetchOptions.headers || {})
  };
  const retriedResponse = await fetchWithRetry(url, {
    ...fetchOptions,
    headers
  }, { maxRetries: retryCountForRequest(fetchOptions, maxRetries) });
  const body = await parseResponseBody(retriedResponse);
  const requestId = graphRequestId(retriedResponse.headers, body);
  rememberGraphRequestId(requestId, { mutation: mutationRequest });
  if (returnResponse) {
    return { body, headers: retriedResponse.headers, status: retriedResponse.status, ok: retriedResponse.ok, graphRequestId: requestId };
  }
  if (!retriedResponse.ok) {
    throw microsoftGraphError(body, retriedResponse);
  }
  return body;
}

async function graphDownloadToFile(path, target, options = {}) {
  const accessToken = await getAccessToken();
  const url = graphUrl(path);
  const response = await fetchWithRetry(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers || {})
    }
  }, { maxRetries: 3 });
  if (!response.ok) {
    const body = await parseResponseBody(response);
    throw microsoftGraphError(body, response);
  }
  await mkdir(dirname(target), { recursive: true, mode: 0o700 });
  const temp = `${target}.part-${process.pid}-${randomUUID()}`;
  let bytesWritten = 0;
  try {
    if (response.body) {
      const counter = new TransformStream({
        transform(chunk, controller) {
          bytesWritten += chunk.byteLength;
          controller.enqueue(chunk);
        }
      });
      await pipeline(Readable.fromWeb(response.body.pipeThrough(counter)), createWriteStream(temp, { flags: "wx", mode: 0o600 }));
    } else {
      const buffer = Buffer.from(await response.arrayBuffer());
      bytesWritten = buffer.length;
      await writePrivateFile(temp, buffer);
    }
    await renameFile(temp, target);
    await hardenPrivateFile(target);
  } catch (error) {
    await rm(temp, { force: true });
    if (options.reserved === true) await rm(target, { force: true });
    throw error;
  }
  return { bytesWritten: bytesWritten || contentLength(response) || 0 };
}

async function graphLimitedBuffer(path, maxBytes, options = {}) {
  const accessToken = await getAccessToken();
  const url = graphUrl(path);
  const limit = Math.max(1, maxBytes);
  const response = await fetchWithRetry(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Range: `bytes=0-${limit}`
    }
  }, { maxRetries: 3 });
  if (!response.ok && response.status !== 206) {
    const body = await parseResponseBody(response);
    throw microsoftGraphError(body, response);
  }

  if (!response.body?.getReader) {
    const buffer = Buffer.from(await response.arrayBuffer());
    return {
      buffer: buffer.subarray(0, limit),
      bytesRead: buffer.length,
      truncated: buffer.length > limit
    };
  }

  const reader = response.body.getReader();
  const chunks = [];
  let bytesRead = 0;
  let truncated = false;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = Buffer.from(value);
      bytesRead += chunk.length;
      const remaining = limit + 1 - chunks.reduce((sum, entry) => sum + entry.length, 0);
      if (remaining > 0) chunks.push(chunk.subarray(0, remaining));
      if (bytesRead > limit) {
        truncated = true;
        await reader.cancel().catch(() => {});
        break;
      }
    }
  } finally {
    reader.releaseLock?.();
  }

  const buffer = Buffer.concat(chunks);
  return {
    buffer: buffer.subarray(0, limit),
    bytesRead,
    truncated: truncated || bytesRead > limit || buffer.length > limit
  };
}

function cleanPath(path = "") {
  return String(path).replace(/^\/+/, "").replace(/\/+$/, "");
}

function clampInteger(value, defaultValue, min, max) {
  const number = Number(value ?? defaultValue);
  if (!Number.isInteger(number)) return defaultValue;
  return Math.min(Math.max(number, min), max);
}

function assertSafeRemotePath(path = "", label = "path") {
  const clean = cleanPath(path);
  for (const segment of clean.split("/").filter(Boolean)) {
    if (segment === "." || segment === ".." || segment.includes("\\")) {
      throw new Error(`${label} contains an unsafe path segment: ${segment}`);
    }
  }
  return clean;
}

function assertSafeItemName(name = "", label = "name") {
  if (typeof name !== "string") throw new Error(`${label} must be a string.`);
  const value = String(name);
  const trimmed = value.trim();
  if (!trimmed) throw new Error(`${label} must not be empty.`);
  if (trimmed === "." || trimmed === "..") throw new Error(`${label} must not be "." or "..".`);
  if (/[\/\\]/.test(value)) throw new Error(`${label} must be a single item name, not a path.`);
  if (/[\u0000-\u001F]/.test(value)) throw new Error(`${label} contains control characters.`);
  return trimmed;
}

function assertAtMostOneSelector(args = {}, label, selectors = []) {
  const provided = selectors.filter((selector) =>
    selector.keys.every((key) => Object.hasOwn(args, key) && args[key] !== undefined && args[key] !== null && args[key] !== "")
  );
  if (provided.length > 1) {
    throw new Error(`${label} must use at most one selector: ${selectors.map((selector) => selector.label).join(" or ")}.`);
  }
}

function pathName(path = "") {
  return basename(cleanPath(path));
}

function encodeDrivePath(path) {
  return assertSafeRemotePath(path).split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

function joinRemotePath(...parts) {
  return assertSafeRemotePath(parts.filter((part) => part !== undefined && part !== null && String(part).trim() !== "").join("/"));
}

function resolvePresetPath(args = {}, options = {}) {
  const {
    pathField = "path",
    presetField = "preset",
    relativeField = "relativePath",
    allowEmpty = true
  } = options;
  const explicitPath = args[pathField];
  const preset = args[presetField];
  if (preset) {
    const presets = pathPresets();
    if (!Object.hasOwn(presets, preset)) {
      throw new Error(`Unknown OneDrive path preset: ${preset}. Available presets: ${Object.keys(presets).join(", ")}`);
    }
    return joinRemotePath(presets[preset], args[relativeField] || "");
  }
  const resolved = assertSafeRemotePath(explicitPath || "", pathField);
  if (!allowEmpty && !resolved) throw new Error(`${pathField} or ${presetField} is required.`);
  return resolved;
}

function itemArgsWithResolvedPath(args = {}) {
  if (args.itemId) return args;
  return { ...args, path: resolvePresetPath(args) };
}

function remotePath(args = {}) {
  if (args.remotePreset && !args.remoteRelativePath) {
    throw new Error("remoteRelativePath is required when remotePreset is used, and must include the destination filename.");
  }
  const resolved = resolvePresetPath(args, {
    pathField: "remotePath",
    presetField: "remotePreset",
    relativeField: "remoteRelativePath",
    allowEmpty: false
  });
  if (args.remotePreset) splitRemotePath(resolved);
  return resolved;
}

function normalizeLocalPathForCompare(path) {
  return resolve(path).replace(/\/+$/g, "");
}

function isLocalOneDriveSyncPath(path) {
  const normalized = normalizeLocalPathForCompare(path);
  return localOneDriveSyncRoots.some((root) => {
    const normalizedRoot = normalizeLocalPathForCompare(root.path);
    return normalized === normalizedRoot
      || normalized.startsWith(`${normalizedRoot}/`)
      || (root.prefix && normalized.startsWith(normalizedRoot));
  });
}

function assertNotLocalOneDriveSyncPath(path, operation, args = {}) {
  if (args.allowLocalOneDriveSyncPath === true) return;
  if (!isLocalOneDriveSyncPath(path)) return;
  throw new Error(`${operation} refuses to use a local OneDrive sync folder by default: ${path}. Use remote OneDrive plugin tools instead, or pass allowLocalOneDriveSyncPath: true only when you explicitly need local sync-folder access.`);
}

async function existingRealPath(path) {
  try {
    return await realpath(path);
  } catch {
    return null;
  }
}

async function realPathForPotentialWrite(path) {
  const target = resolve(path);
  const existingTarget = await existingRealPath(target);
  if (existingTarget) return existingTarget;

  let parent = dirname(target);
  while (parent && parent !== dirname(parent)) {
    const existingParent = await existingRealPath(parent);
    if (existingParent) return join(existingParent, relative(parent, target));
    parent = dirname(parent);
  }
  return target;
}

async function assertNotLocalOneDriveSyncPathForRead(path, operation, args = {}) {
  assertNotLocalOneDriveSyncPath(path, operation, args);
  const real = await existingRealPath(path);
  if (real) assertNotLocalOneDriveSyncPath(real, operation, args);
}

async function assertNotLocalOneDriveSyncPathForWrite(path, operation, args = {}) {
  assertNotLocalOneDriveSyncPath(path, operation, args);
  const real = await realPathForPotentialWrite(path);
  assertNotLocalOneDriveSyncPath(real, operation, args);
}

function formatDriveItem(item, format = "compact") {
  const simplified = simplifyItem(item);
  return formatSimplifiedItem(simplified, format);
}

function formatSimplifiedItem(simplified, format = "compact") {
  if (!simplified || format === "full") return simplified;
  return {
    id: simplified.id,
    name: simplified.name,
    remotePath: simplified.remotePath,
    type: simplified.folder ? "folder" : simplified.file ? "file" : "item",
    size: simplified.size,
    lastModifiedDateTime: simplified.lastModifiedDateTime,
    webUrl: simplified.webUrl
  };
}

function itemBase(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (resolved.itemId) return `/me/drive/items/${encodeURIComponent(resolved.itemId)}`;
  if (!resolved.path || cleanPath(resolved.path) === "") return "/me/drive/root";
  return `/me/drive/root:/${encodeDrivePath(resolved.path)}:`;
}

function itemIdBase(itemId) {
  return `/me/drive/items/${encodeURIComponent(itemId)}`;
}

function requireNonRootTarget(args = {}, operation) {
  const resolved = itemArgsWithResolvedPath(args);
  if (!resolved.itemId && (!resolved.path || cleanPath(resolved.path) === "")) {
    throw new Error(`${operation} requires a non-root path or itemId. Refusing to operate on the OneDrive root.`);
  }
}

function assertExpectedItem(rawItem, args = {}, operation) {
  if (args.expectedId && rawItem.id !== args.expectedId) {
    throw new Error(`${operation} expected item ID ${args.expectedId}, but resolved ${rawItem.id}. Refusing to continue.`);
  }
  if (args.expectedName && rawItem.name !== args.expectedName) {
    throw new Error(`${operation} expected item named ${args.expectedName}, but resolved ${rawItem.name}. Refusing to continue.`);
  }
}

function hasExpectedIdentity(args = {}) {
  return Boolean(args.expectedName || args.expectedId);
}

function stablePreviewValue(value) {
  if (Array.isArray(value)) return value.map(stablePreviewValue);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, stablePreviewValue(value[key])])
      .filter(([, child]) => child !== undefined)
  );
}

function previewProofHash(tool, proof = {}) {
  return createHash("sha256")
    .update(tool)
    .update("\n")
    .update(JSON.stringify(stablePreviewValue(proof)))
    .digest("hex");
}

function cleanupPreviewTokens(now = Date.now()) {
  for (const [token, entry] of previewTokens.entries()) {
    if (!entry || entry.expiresAt <= now) previewTokens.delete(token);
  }
}

function issuePreviewToken(tool, proof = {}) {
  cleanupPreviewTokens();
  const token = randomUUID();
  const expiresAt = Date.now() + previewTokenTtlMs;
  previewTokens.set(token, {
    tool,
    proofHash: previewProofHash(tool, proof),
    expiresAt
  });
  return {
    previewToken: token,
    previewTokenExpiresAt: new Date(expiresAt).toISOString()
  };
}

function previewWithToken(preview, tool, proof = {}) {
  if (preview?.dryRun !== true) return preview;
  return {
    ...preview,
    ...issuePreviewToken(tool, proof)
  };
}

function consumePreviewToken(tool, proof = {}, token) {
  cleanupPreviewTokens();
  if (!token) return { ok: false, reason: "missing" };
  const entry = previewTokens.get(token);
  if (!entry) return { ok: false, reason: "not_found_or_expired" };
  if (entry.tool !== tool || entry.proofHash !== previewProofHash(tool, proof)) {
    return { ok: false, reason: "mismatch" };
  }
  previewTokens.delete(token);
  return { ok: true };
}

function previewTokenRequiredResult(preview, tool, proof, token, requiredField) {
  const result = consumePreviewToken(tool, proof, token);
  if (result.ok) return null;
  return {
    ...preview,
    dryRun: false,
    confirmed: true,
    previewTokenRequired: true,
    previewTokenStatus: result.reason,
    [requiredField]: "Run a dry-run preview for this exact operation and pass the returned previewToken with dryRun: false and confirmed: true."
  };
}

function batchMutationWarnings(extra = []) {
  return [partialBatchMutationWarning, ...extra.filter(Boolean)];
}

function itemAuditSummary(item) {
  const simplified = item?.remotePath !== undefined ? item : simplifyItem(item);
  if (!simplified) return null;
  return {
    id: simplified.id,
    name: simplified.name,
    remotePath: simplified.remotePath,
    type: simplified.folder ? "folder" : simplified.file ? "file" : "item",
    size: simplified.size,
    lastModifiedDateTime: simplified.lastModifiedDateTime
  };
}

function itemMutationBase(rawItem) {
  if (!rawItem?.id) throw new Error("Resolved item is missing an ID; refusing live mutation.");
  return itemIdBase(rawItem.id);
}

function mutationMatchHeaders(rawItem) {
  return rawItem?.eTag ? { "If-Match": rawItem.eTag } : {};
}

function permissionAuditSummary(permission = {}) {
  return {
    id: permission.id,
    roles: permission.roles,
    link: permission.link ? {
      type: permission.link.type,
      scope: permission.link.scope,
      preventsDownload: permission.link.preventsDownload
    } : undefined,
    grantedToPresent: Boolean(permission.grantedTo || permission.grantedToV2),
    grantedToIdentityCount: Array.isArray(permission.grantedToIdentities) ? permission.grantedToIdentities.length : undefined,
    grantedToIdentityV2Count: Array.isArray(permission.grantedToIdentitiesV2) ? permission.grantedToIdentitiesV2.length : undefined,
    invitationPresent: Boolean(permission.invitation),
    invitationSignInRequired: permission.invitation?.signInRequired,
    inheritedFrom: permission.inheritedFrom ? itemAuditSummary(permission.inheritedFrom) : undefined,
    expirationDateTime: permission.expirationDateTime,
    hasPassword: permission.hasPassword
  };
}

function permissionDiffAuditSummary(diff = {}) {
  return {
    added: (diff.added || []).map(permissionAuditSummary),
    removed: (diff.removed || []).map(permissionAuditSummary),
    unchangedCount: diff.unchangedCount || 0,
    beforeCount: diff.beforeCount || 0,
    afterCount: diff.afterCount || 0
  };
}

function safeErrorInfo(error) {
  return {
    message: redactAuditText(error?.message || String(error)),
    graphRequestId: error?.graphRequestId || currentGraphRequestId({ mutation: true }) || currentGraphRequestId() || undefined,
    graphStatus: error?.graphStatus
  };
}

function redactAuditText(text = "") {
  return String(text)
    .replace(/https?:\/\/[^\s")]+/gi, "[redacted-url]")
    .replace(/\bBearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [redacted]")
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, "[redacted-email]")
    .replace(/\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi, "[redacted-id]")
    .replace(/\b(access[_-]?token|refresh[_-]?token|id[_-]?token|secret)\s*[:=]\s*([^\s,;)]+)/gi, "$1=[redacted]")
    .replace(/\b(email|alias|objectId|recipient)\s*[:=]\s*([^\s,;)]+)/gi, "$1=[redacted]");
}

function safeToolErrorMessage(error) {
  return redactAuditText(error?.message || String(error));
}

function recordLocalWarning(operation, error) {
  const store = toolCallContext.getStore();
  if (!store) return;
  store.localWarnings.push({ operation, error: safeToolErrorMessage(error) });
}

async function bestEffortLocalWrite(operation, fn) {
  try {
    return await fn();
  } catch (error) {
    recordLocalWarning(operation, error);
    return null;
  }
}

function sanitizeAuditValue(value) {
  if (Array.isArray(value)) return value.map(sanitizeAuditValue);
  if (!value || typeof value !== "object") return value;
  const clean = {};
  const forbidden = new Set([
    "access_token", "refresh_token", "id_token", "Authorization", "authorization",
    "content", "body", "uploadUrl", "monitorUrl", "resourceLocation", "webUrl",
    "password", "recipients", "email", "alias", "objectId"
  ]);
  for (const [key, entry] of Object.entries(value)) {
    if (forbidden.has(key)) continue;
    clean[key] = sanitizeAuditValue(entry);
  }
  return clean;
}

async function writeMutationAudit(tool, entry) {
  const record = sanitizeAuditValue({
    timestamp: new Date().toISOString(),
    tool,
    ...entry,
    graphRequestId: entry.graphRequestId || currentGraphRequestId({ mutation: true }) || currentGraphRequestId() || undefined
  });
  return await bestEffortLocalWrite("mutation audit write", async () => {
    await withFileLock(auditLockPath, async () => {
      await ensurePrivateDirectory(auditRoot);
      await appendFile(auditPath, `${JSON.stringify(record)}\n`, { encoding: "utf8", mode: 0o600 });
      await hardenPrivateFile(auditPath);
    });
    return record;
  });
}

async function auditRecent(args = {}) {
  let text = "";
  await withFileLock(auditLockPath, async () => {
    await ensurePrivateDirectory(auditRoot);
    try {
      text = await readFile(auditPath, "utf8");
      await hardenPrivateFile(auditPath);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  });
  const limit = clampInteger(args.limit, 50, 1, 500);
  const since = args.since ? Date.parse(args.since) : null;
  const until = args.until ? Date.parse(args.until) : null;
  if (args.since && Number.isNaN(since)) throw new Error("since must be an ISO timestamp.");
  if (args.until && Number.isNaN(until)) throw new Error("until must be an ISO timestamp.");
  const pathNeedle = args.pathContains ? String(args.pathContains).toLowerCase() : "";
  const entryPathText = (entry) => JSON.stringify({
    target: entry.target,
    before: entry.before,
    after: entry.after
  }).toLowerCase();
  let entries = text.trim()
    ? text.trim().split("\n").map((line) => {
        try {
          return JSON.parse(line);
        } catch {
          return { malformed: true, line };
        }
      })
    : [];
  entries = entries.filter((entry) => {
    if (args.tool && entry.tool !== args.tool) return false;
    if (args.status && entry.status !== args.status) return false;
    const timestamp = entry.timestamp ? Date.parse(entry.timestamp) : null;
    if (since !== null && (!timestamp || timestamp < since)) return false;
    if (until !== null && (!timestamp || timestamp > until)) return false;
    if (pathNeedle && !entryPathText(entry).includes(pathNeedle)) return false;
    return true;
  });
  if (args.newestFirst !== false) entries.reverse();
  entries = entries.slice(0, limit);
  return {
    auditPath,
    count: entries.length,
    filters: {
      tool: args.tool || null,
      status: args.status || null,
      pathContains: args.pathContains || null,
      since: args.since || null,
      until: args.until || null,
      newestFirst: args.newestFirst !== false
    },
    entries
  };
}

async function auditExport(args = {}) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const preferredTarget = args.localPath ? resolve(args.localPath) : join(auditRoot, `export-${stamp}-${randomUUID()}.jsonl`);
  await assertNotLocalOneDriveSyncPathForWrite(preferredTarget, "Audit export", args);
  const reservation = await reserveLocalDestination(preferredTarget, {
    overwrite: args.overwrite === true,
    allowAlternate: !args.localPath
  });
  const target = reservation.path;
  try {
    await withFileLock(auditLockPath, async () => {
      try {
        await copyFile(auditPath, target);
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
        await writePrivateFile(target, "");
      }
    });
    await hardenPrivateFile(target);
    const written = await stat(target);
    return { auditPath, localPath: target, bytesWritten: written.size };
  } catch (error) {
    if (reservation.reserved) await rm(target, { force: true });
    throw error;
  }
}

async function auditClear(args = {}) {
  if (args.confirmed !== true) {
    return {
      confirmed: false,
      auditPath,
      requiredToClear: "Set confirmed: true after explicit user confirmation to clear the local mutation audit log."
    };
  }
  await withFileLock(auditLockPath, async () => await rm(auditPath, { force: true }));
  return { confirmed: true, auditPath, cleared: true };
}

function childrenPath(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (resolved.itemId) return `/me/drive/items/${encodeURIComponent(resolved.itemId)}/children`;
  if (!resolved.path || cleanPath(resolved.path) === "") return "/me/drive/root/children";
  return `/me/drive/root:/${encodeDrivePath(resolved.path)}:/children`;
}

function contentPath(args = {}) {
  return `${itemBase(args)}/content`;
}

function uploadPath(remotePath, conflictBehavior = "fail") {
  const behavior = encodeURIComponent(conflictBehavior);
  return `/me/drive/root:/${encodeDrivePath(remotePath)}:/content?@microsoft.graph.conflictBehavior=${behavior}`;
}

function splitRemotePath(remotePath) {
  const clean = assertSafeRemotePath(remotePath, "remotePath");
  if (!clean) throw new Error("remotePath must include a filename.");
  const parts = clean.split("/").filter(Boolean);
  const name = parts.pop();
  return { parentPath: parts.join("/"), name };
}

async function uploadSessionTarget(remotePath) {
  const { parentPath, name } = splitRemotePath(remotePath);
  const parent = await resolveDestinationParent({ destinationParentPath: parentPath });
  return {
    endpoints: [
      `/me/drive/items/${parent.id}:/${encodeURIComponent(name)}:/createUploadSession`,
      `/me/drive/root:/${encodeDrivePath(remotePath)}:/createUploadSession`
    ],
    name
  };
}

function simplifyItem(item) {
  if (!item) return item;
  const source = item.remoteItem ? {
    ...item.remoteItem,
    id: item.remoteItem.id || item.id,
    name: item.remoteItem.name || item.name,
    webUrl: item.remoteItem.webUrl || item.webUrl,
    parentReference: item.remoteItem.parentReference || item.parentReference
  } : item;
  return {
    id: source.id,
    name: source.name,
    remotePath: itemRemotePath(source),
    path: source.parentReference?.path,
    parentId: source.parentReference?.id,
    driveId: source.parentReference?.driveId,
    webUrl: source.webUrl,
    size: source.size,
    createdDateTime: source.createdDateTime,
    lastModifiedDateTime: source.lastModifiedDateTime,
    eTag: source.eTag,
    cTag: source.cTag,
    deleted: source.deleted,
    folder: source.folder ? { childCount: source.folder.childCount } : undefined,
    file: source.file ? {
      mimeType: source.file.mimeType,
      ...(source.file.hashes !== undefined ? { hashes: source.file.hashes } : {})
    } : undefined
  };
}

function itemRemotePath(item) {
  if (!item?.name) return undefined;
  if (!item.parentReference && !item.root) return undefined;
  const parentPath = item.parentReference?.path;
  if (!parentPath && item.parentReference?.id) return undefined;
  const rootMatch = String(parentPath || "").match(/^\/(?:drive|drives\/[^/]+)\/root:(.*)$/);
  const parentRemotePath = rootMatch
    ? decodeGraphPath(rootMatch[1].replace(/^\/+|\/+$/g, ""))
    : "";
  return [parentRemotePath, item.name].filter(Boolean).join("/");
}

function decodeGraphPath(path = "") {
  try {
    return decodeURIComponent(path);
  } catch {
    return path;
  }
}

async function list(args = {}) {
  const params = new URLSearchParams();
  params.set("$top", String(clampInteger(args.limit, 100, 1, 200)));
  params.set("$select", args.select || defaultSelect);
  const result = await graph(`${childrenPath(args)}?${params.toString()}`);
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(result.value || []));
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), nextLink: result["@odata.nextLink"] || null };
}

function isDeltaCursor(value) {
  if (!value) return false;
  try {
    const base = new URL(graphBaseUrl());
    const target = new URL(String(value), base);
    const basePath = base.pathname.replace(/\/+$/, "");
    const trustedPath = target.origin === base.origin
      && (target.pathname === basePath || target.pathname.startsWith(`${basePath}/`));
    return trustedPath && target.pathname.split("/").some((segment) => segment === "delta" || segment.startsWith("delta("));
  } catch {
    return false;
  }
}

async function collectPages(firstPath, maxItems, format = "compact", formatter = formatDriveItem, options = {}) {
  return await withMetadataCacheBatch(async () => {
    const items = [];
    let nextPath = firstPath;
    let nextLink = null;
    let deltaLink = null;
    let truncated = false;
    const seenPages = new Set();
    let pagesFetched = 0;
    let unsafePageTruncation = false;
    const pendingCacheItems = [];
    let pendingCursorMetadata = {};
    const maxPages = Math.max(1, Math.ceil(maxItems / 1) + 100);
    while (nextPath && items.length < maxItems) {
      if (seenPages.has(nextPath)) throw new Error(`Microsoft Graph pagination cycle detected at ${safeDisplayPath(nextPath)}.`);
      if (pagesFetched >= maxPages) throw new Error(`Microsoft Graph pagination exceeded ${maxPages} pages before reaching the item limit.`);
      seenPages.add(nextPath);
      pagesFetched += 1;
      let result;
      try {
        result = await graph(nextPath);
      } catch (error) {
        error.graphPagesFetched = pagesFetched;
        throw error;
      }
      const pageItems = result.value || [];
      const remaining = maxItems - items.length;
      const acceptedItems = pageItems.slice(0, remaining);
      const pageTruncated = pageItems.length > remaining;
      const pageNextLink = result["@odata.nextLink"] || null;
      const pageDeltaLink = result["@odata.deltaLink"] || null;
      const cursorMetadata = options.persistDeltaCursor === true && !pageTruncated && (pageDeltaLink || pageNextLink) ? {
        deltaLink: pageDeltaLink || undefined,
        deltaNextLink: pageNextLink && !pageDeltaLink ? pageNextLink : undefined,
        deltaTarget: options.deltaTarget
      } : {};
      const cacheableItems = options.prepareCacheItems
        ? await options.prepareCacheItems(acceptedItems)
        : acceptedItems;
      if (options.cacheResults !== false) pendingCacheItems.push(...cacheableItems);
      if (Object.keys(cursorMetadata).length) pendingCursorMetadata = cursorMetadata;
      items.push(...acceptedItems.map((item) => formatter(item, format)));
      if (pageTruncated) {
        unsafePageTruncation = true;
        nextLink = null;
      } else {
        nextLink = pageNextLink;
      }
      deltaLink = pageTruncated ? null : pageDeltaLink;
      nextPath = !pageTruncated && nextLink && items.length < maxItems ? nextLink : null;
      truncated = truncated || pageTruncated || (Boolean(nextLink) && items.length >= maxItems);
    }
    if (pendingCacheItems.length || Object.keys(pendingCursorMetadata).length) {
      const deduplicatedCacheItems = [...new Map(pendingCacheItems.filter((item) => item?.id).map((item) => [item.id, item])).values()];
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(deduplicatedCacheItems, pendingCursorMetadata));
    }
    return {
      items,
      nextLink,
      deltaLink,
      truncated,
      unsafePageTruncation,
      pagesFetched,
      count: items.length
    };
  });
}

function safeDisplayPath(value = "") {
  try {
    const url = new URL(value, graphBaseUrl());
    return `${url.origin}${url.pathname}`;
  } catch {
    return String(value).split("?")[0];
  }
}

async function listAll(args = {}) {
  const maxItems = clampInteger(args.maxItems, 1000, 1, 5000);
  const params = new URLSearchParams();
  params.set("$top", String(Math.min(clampInteger(args.pageSize, 200, 1, 200), maxItems)));
  params.set("$select", args.select || defaultSelect);
  return await collectPages(`${childrenPath(args)}?${params.toString()}`, maxItems, args.format, formatDriveItem, {
    cacheResults: args.cacheResults !== false
  });
}

function normalizeExtensions(extensions = []) {
  return new Set(
    extensions
      .map((extension) => String(extension || "").trim().toLowerCase())
      .filter(Boolean)
      .map((extension) => extension.startsWith(".") ? extension : `.${extension}`)
  );
}

function scanItemMatches(item, args = {}, extensionFilter = new Set()) {
  const isFolder = Boolean(item.folder);
  const isFile = Boolean(item.file);
  if (isFile && args.includeFiles === false) return false;
  if (isFolder && args.includeFolders === false) return false;
  if (!isFile && !isFolder && args.includeFiles === false && args.includeFolders === false) return false;

  if (args.nameContains) {
    const needle = String(args.nameContains).toLowerCase();
    if (!String(item.name || "").toLowerCase().includes(needle)) return false;
  }

  if (extensionFilter.size && isFile) {
    const extension = extname(item.name || "").toLowerCase();
    if (!extensionFilter.has(extension)) return false;
  }

  if (extensionFilter.size && isFolder && args.includeFolders === false) return false;
  return true;
}

async function resolveScanRoot(args = {}) {
  if (args.itemId) {
    const folder = await getRawInfo({ itemId: args.itemId, cacheResults: args.cacheResults });
    if (!folder.folder && !folder.root) throw new Error(`Scan target is not a folder: ${folder.name}`);
    return { id: folder.id, name: folder.name || "root", remotePath: itemRemotePath(folder) || "", target: `itemId:${folder.id}` };
  }

  const resolvedPath = resolvePresetPath(args);
  if (resolvedPath) {
    const folder = await getRawInfo({ path: resolvedPath, cacheResults: args.cacheResults });
    if (!folder.folder && !folder.root) throw new Error(`Scan target is not a folder: ${folder.name}`);
    return { id: folder.id, name: folder.name || resolvedPath, remotePath: resolvedPath, target: resolvedPath };
  }

  const root = await graph("/me/drive/root");
  return { id: root.id, name: "root", remotePath: "", target: "root" };
}

async function scan(args = {}) {
  return await withMetadataCacheBatch(async () => {
    const pageSize = clampInteger(args.pageSize, 200, 1, 200);
    const maxItems = clampInteger(args.maxItems, 10000, 1, 50000);
    const maxResults = clampInteger(args.maxResults, 500, 1, 5000);
    const maxDepth = clampInteger(args.maxDepth, 25, 0, 50);
    const maxFolders = clampInteger(args.maxFolders, 1000, 1, 10000);
    const extensionFilter = normalizeExtensions(args.extensions || []);
    const skipFolderIds = new Set((args._skipFolderIds || []).filter(Boolean));
    const root = args._resolvedRoot || await resolveScanRoot(args);
    const params = new URLSearchParams();
    params.set("$top", String(pageSize));
    params.set("$select", args.select || defaultSelect);

    const queue = [{ id: root.id, name: root.name, remotePath: root.remotePath, depth: 0 }];
    const results = [];
    const pendingCacheItems = [];
    const counters = {
      itemsScanned: 0,
      filesScanned: 0,
      foldersScanned: 0,
      foldersVisited: 0,
      foldersSkipped: 0,
      matched: 0
    };
    let truncatedReason = null;

    while (queue.length) {
      if (counters.itemsScanned >= maxItems) {
        truncatedReason = "maxItems";
        break;
      }
      if (counters.foldersVisited >= maxFolders) {
        truncatedReason = "maxFolders";
        break;
      }

      const folder = queue.shift();
      counters.foldersVisited += 1;
      let nextPath = `/me/drive/items/${encodeURIComponent(folder.id)}/children?${params.toString()}`;
      const seenPages = new Set();
      let pagesFetched = 0;
      const maxPagesPerFolder = Math.max(1, maxItems + 100);

      while (nextPath) {
        if (counters.itemsScanned >= maxItems) {
          truncatedReason = "maxItems";
          break;
        }
        if (seenPages.has(nextPath)) {
          throw new Error(`Microsoft Graph pagination cycle detected while scanning ${folder.remotePath || folder.name || folder.id}.`);
        }
        if (pagesFetched >= maxPagesPerFolder) {
          throw new Error(`Microsoft Graph pagination exceeded ${maxPagesPerFolder} pages while scanning ${folder.remotePath || folder.name || folder.id}.`);
        }
        seenPages.add(nextPath);
        pagesFetched += 1;
        const page = await graph(nextPath);
        const cacheableItems = [];
        for (const item of page.value || []) {
          if (counters.itemsScanned >= maxItems) {
            truncatedReason = "maxItems";
            break;
          }
          cacheableItems.push(item);
          counters.itemsScanned += 1;
          if (item.file) counters.filesScanned += 1;
          if (item.folder) counters.foldersScanned += 1;
          if (typeof args.onItem === "function") await args.onItem(item);

          if (scanItemMatches(item, args, extensionFilter)) {
            counters.matched += 1;
            if (results.length < maxResults) {
              results.push(formatDriveItem(item, args.format));
            }
            if (args.stopAfterResults === true && results.length >= maxResults) {
              truncatedReason = "maxResults";
              break;
            }
          }

          if (item.folder && folder.depth < maxDepth) {
            if (skipFolderIds.has(item.id)) {
              counters.foldersSkipped += 1;
            } else {
              queue.push({
                id: item.id,
                name: item.name,
                remotePath: itemRemotePath(item),
                depth: folder.depth + 1
              });
            }
          }
        }
        if (args.cacheResults !== false) pendingCacheItems.push(...cacheableItems);
        if (truncatedReason) break;
        nextPath = page["@odata.nextLink"] || null;
      }

      if (truncatedReason) break;
    }

    if (!truncatedReason && queue.length) truncatedReason = "queueRemaining";
    if (args.cacheResults !== false && pendingCacheItems.length) {
      const deduplicatedCacheItems = [...new Map(pendingCacheItems.filter((item) => item?.id).map((item) => [item.id, item])).values()];
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(deduplicatedCacheItems));
    }
    return {
      root,
      filters: {
        nameContains: args.nameContains || null,
        extensions: [...extensionFilter],
        includeFiles: args.includeFiles !== false,
        includeFolders: args.includeFolders !== false,
        maxDepth,
        maxItems,
        maxFolders,
        maxResults
      },
      summary: {
        ...counters,
        returned: results.length,
        resultTruncated: counters.matched > results.length,
        traversalTruncated: Boolean(truncatedReason),
        truncatedReason,
        foldersQueued: queue.length
      },
      items: results,
      note: truncatedReason
        ? `Scan stopped at ${truncatedReason}. Increase the relevant cap or narrow the scan.`
        : "Recursive scan completed within the requested caps."
    };
  });
}

async function search(args = {}) {
  const escaped = String(args.query).replace(/'/g, "''");
  const params = new URLSearchParams();
  params.set("$top", String(clampInteger(args.limit, 50, 1, 200)));
  params.set("$select", defaultSelect);
  const result = await graph(`/me/drive/root/search(q='${encodeURIComponent(escaped)}')?${params.toString()}`);
  if (args.cacheResults !== false) {
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(result.value || []));
  }
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), nextLink: result["@odata.nextLink"] || null };
}

async function searchAll(args = {}) {
  const escaped = String(args.query).replace(/'/g, "''");
  const maxItems = clampInteger(args.maxItems, 1000, 1, 5000);
  const params = new URLSearchParams();
  params.set("$top", String(Math.min(clampInteger(args.pageSize, 200, 1, 200), maxItems)));
  params.set("$select", defaultSelect);
  return await collectPages(
    `/me/drive/root/search(q='${encodeURIComponent(escaped)}')?${params.toString()}`,
    maxItems,
    args.format,
    formatDriveItem,
    { cacheResults: args.cacheResults !== false }
  );
}

const findStopWords = new Set([
  "a", "an", "and", "by", "can", "could", "find", "for", "from", "get", "i", "in", "is", "it",
  "locate", "me", "my", "named", "of", "on", "or", "please", "show", "the", "to", "where", "with"
]);
const findGenericWords = new Set([
  "called", "file", "files", "folder", "folders", "named", "document", "documents", "summary",
  "codex", "onedrive", "plugin", "test"
]);
const findKindHints = [
  {
    kind: "presentation",
    words: ["deck", "decks", "presentation", "presentations", "powerpoint", "ppt", "pptx", "slides", "slideshow"],
    extensions: [".pptx", ".ppt", ".pptm", ".ppsx", ".odp"]
  },
  {
    kind: "spreadsheet",
    words: ["spreadsheet", "spreadsheets", "sheet", "sheets", "excel", "xlsx", "xls", "csv", "workbook"],
    extensions: [".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods"]
  },
  {
    kind: "word",
    words: ["doc", "docs", "docx", "word", "letter", "memo"],
    extensions: [".docx", ".doc", ".docm", ".rtf", ".odt"]
  },
  {
    kind: "pdf",
    words: ["pdf"],
    extensions: [".pdf"]
  },
  {
    kind: "image",
    words: ["image", "images", "photo", "photos", "picture", "pictures", "screenshot", "screenshots", "png", "jpg", "jpeg"],
    extensions: [".png", ".jpg", ".jpeg", ".gif", ".webp", ".heic", ".tif", ".tiff"]
  }
];

function normalizeFindText(value = "") {
  return String(value)
    .toLowerCase()
    .replace(/[_]+/g, " ")
    .replace(/[^\p{L}\p{N}.-]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function findTokens(value = {}) {
  const normalized = normalizeFindText(value);
  const rawTokens = normalized.match(/[\p{L}\p{N}]+(?:[-.][\p{L}\p{N}]+)*/gu) || [];
  const expanded = [];
  for (const token of rawTokens) {
    expanded.push(token);
    if (/[-.]/.test(token)) expanded.push(...token.split(/[-.]/).filter(Boolean));
  }
  return [...new Set(expanded.filter((token) => token.length > 1 && !findStopWords.has(token)))];
}

function findDateTokens(query) {
  return [...new Set(String(query).match(/\b\d{4}[-_. ]\d{1,2}[-_. ]\d{1,2}\b/g) || [])]
    .map((token) => token.trim().replace(/[_. ]/g, "-"));
}

function findImportantTokens(query) {
  return findTokens(query).filter((token) => !findGenericWords.has(token));
}

function buildFindSearchTerms(query, maxSearchTerms = 8) {
  const terms = [];
  const add = (term) => {
    const clean = String(term || "").replace(/\s+/g, " ").trim();
    if (clean && !terms.some((existing) => existing.toLowerCase() === clean.toLowerCase())) terms.push(clean);
  };
  const dateTokens = findDateTokens(query);
  const dateParts = new Set(dateTokens.flatMap((token) => [token, ...token.split("-")]));
  const important = findImportantTokens(query).filter((token) => !dateParts.has(token));

  add(query);
  for (const dateToken of dateTokens) add(dateToken);
  if (important.length) add(important.join(" "));
  if (important.length >= 3) add(important.slice(-3).join(" "));
  if (important.length >= 2) {
    for (let index = 0; index < important.length - 1; index += 1) {
      add(`${important[index]} ${important[index + 1]}`);
    }
  }
  for (const token of [...important].sort((a, b) => b.length - a.length)) {
    if (token.length >= 4) add(token);
  }

  return terms.slice(0, Math.min(maxSearchTerms, 12));
}

function inferFindExtensions(query, explicitExtensions = []) {
  const explicit = normalizeExtensions(explicitExtensions);
  const inferred = new Set();
  const strictInferred = new Set();
  const tokens = new Set(findTokens(query));
  const normalized = normalizeFindText(query);
  const matchedKinds = [];
  for (const hint of findKindHints) {
    if (hint.words.some((word) => tokens.has(word) || normalized.includes(word))) {
      matchedKinds.push(hint.kind);
      for (const extension of hint.extensions) inferred.add(extension);
      const strictMatch = hint.words.some((word) => tokens.has(word) && (word.startsWith(".") || hint.extensions.includes(`.${word}`) || ["pdf", "ppt", "pptx", "powerpoint", "excel", "xlsx", "xls", "csv", "doc", "docx", "word", "png", "jpg", "jpeg"].includes(word)));
      if (strictMatch) {
        for (const extension of hint.extensions) strictInferred.add(extension);
      }
    }
  }
  return {
    explicit,
    inferred,
    strictInferred,
    effectiveForScan: explicit.size ? explicit : strictInferred,
    matchedKinds
  };
}

function findItemExtension(item = {}) {
  return extname(item.name || "").toLowerCase();
}

function findItemType(item = {}) {
  return item.folder ? "folder" : item.file ? "file" : "item";
}

function findCandidateKey(item = {}) {
  return item.id || item.remotePath || `${item.name || "unnamed"}:${item.size || 0}:${item.lastModifiedDateTime || ""}`;
}

function shouldIncludeFindItem(item, args = {}, extensionInfo = {}) {
  const type = findItemType(item);
  if (type === "folder" && args.includeFolders !== true) return false;
  if (extensionInfo.explicit?.size && type === "file" && !extensionInfo.explicit.has(findItemExtension(item))) return false;
  return true;
}

function scoreFindCandidate(item, context = {}) {
  const queryText = normalizeFindText(context.query);
  const queryTokens = findImportantTokens(context.query);
  const dateTokens = findDateTokens(context.query);
  const nameText = normalizeFindText(item.name || "");
  const pathText = normalizeFindText(item.remotePath || item.path || "");
  const nameTokenSet = new Set(findTokens(item.name || ""));
  const pathTokenSet = new Set(findTokens(item.remotePath || item.path || ""));
  const extension = findItemExtension(item);
  const reasons = [];
  let score = 0;
  let matchedTokens = 0;

  if (context.source === "exactPath") {
    score += 100;
    reasons.push("exact path");
  } else if (context.source === "search") {
    score += 18;
    reasons.push(context.canonicalSearch === true
      ? "canonical Graph content/metadata match"
      : `Graph search: ${context.term}`);
  } else if (context.source === "scan") {
    score += 12;
    reasons.push(`scan: ${context.folder || "root"}`);
  } else if (context.source === "cache") {
    score += 10;
    reasons.push("metadata cache");
  } else if (context.source === "metadataConfirm") {
    score += 12;
    reasons.push("metadata cache confirmed live");
  } else if (context.source === "contentIndex") {
    score += Math.min(55, Math.max(20, context.contentScore || 20));
    matchedTokens += context.contentMatchedTokens || 0;
    reasons.push("indexed content");
    if (context.contentReason) reasons.push(context.contentReason);
  }

  if (queryText && nameText === queryText) {
    score += 90;
    reasons.push("exact filename");
  } else if (queryText && nameText.includes(queryText)) {
    score += 65;
    reasons.push("filename contains full query");
  } else if (queryText && pathText.includes(queryText)) {
    score += 35;
    reasons.push("path contains full query");
  }

  for (const token of queryTokens) {
    if (nameTokenSet.has(token) || nameText.includes(token)) {
      matchedTokens += 1;
      score += token.length >= 5 ? 9 : 5;
    } else if (pathTokenSet.has(token) || pathText.includes(token)) {
      matchedTokens += 1;
      score += token.length >= 5 ? 4 : 2;
    }
  }

  if (queryTokens.length) {
    const coverage = matchedTokens / queryTokens.length;
    score += Math.round(coverage * 45);
    if (coverage >= 0.85) reasons.push("strong token match");
    else if (coverage >= 0.5) reasons.push("partial token match");
  }

  for (const dateToken of dateTokens) {
    const compact = dateToken.replace(/-/g, "");
    if (nameText.includes(dateToken) || pathText.includes(dateToken) || nameText.replace(/[-_. ]/g, "").includes(compact)) {
      score += 18;
      reasons.push(`date match: ${dateToken}`);
    }
  }

  if (context.extensionInfo?.explicit?.size && context.extensionInfo.explicit.has(extension)) {
    score += 35;
    reasons.push(`requested extension: ${extension}`);
  } else if (context.extensionInfo?.inferred?.size && context.extensionInfo.inferred.has(extension)) {
    score += 24;
    reasons.push(`likely file type: ${extension}`);
  }

  if (findItemType(item) === "folder") score -= 15;
  if (item.remotePath && context.scoringFolderHints?.some((hint) => {
    const cleanHint = cleanPath(hint).toLowerCase();
    const cleanItemPath = cleanPath(item.remotePath).toLowerCase();
    return cleanHint && (cleanItemPath === cleanHint || cleanItemPath.startsWith(`${cleanHint}/`));
  })) {
    score += 8;
    reasons.push("folder hint");
  }

  return {
    score: Math.max(0, Math.round(score)),
    matchedTokens,
    reasons: [...new Set(reasons)].slice(0, 6)
  };
}

function defaultFindFolderHints(query = "") {
  const normalized = normalizeFindText(query);
  const hints = [
    "Personal/Documents",
    "Documents",
    "Personal",
    "Microsoft Copilot Chat Files",
    "Desktop"
  ];
  if (normalized.includes("screenshot") || normalized.includes("photo") || normalized.includes("picture")) {
    hints.unshift("Pictures/Screenshots", "Pictures");
  }
  if (normalized.includes("health") || normalized.includes("eval") || normalized.includes("evaluation")) {
    hints.unshift("Personal/Documents/Health");
  }
  return [...new Set(hints)];
}

function normalizeFolderHintKey(hint = "") {
  return cleanPath(hint).toLowerCase();
}

function pruneFolderHints(hints = []) {
  const pruned = [];
  const seen = new Set();
  for (const hint of hints) {
    const clean = cleanPath(hint || "");
    const key = normalizeFolderHintKey(clean);
    if (seen.has(key)) continue;
    seen.add(key);
    if (!key) {
      pruned.push(clean);
      continue;
    }
    const overlapsEarlierHint = pruned.some((earlier) => {
      const earlierKey = normalizeFolderHintKey(earlier);
      if (!earlierKey) return false;
      return key.startsWith(`${earlierKey}/`) || earlierKey.startsWith(`${key}/`);
    });
    if (!overlapsEarlierHint) pruned.push(clean);
  }
  return pruned;
}

function findScanNeedle(query = "") {
  const tokens = findImportantTokens(query)
    .filter((token) => !["deck", "presentation", "powerpoint", "spreadsheet", "sheet", "pdf", "doc", "word"].includes(token))
    .sort((a, b) => b.length - a.length);
  return tokens.find((token) => token.length >= 4) || findDateTokens(query)[0] || tokens[0] || "";
}

function scanNeedleIsWeak(needle = "", query = "") {
  if (!needle) return true;
  const important = findImportantTokens(query);
  if (!important.includes(needle)) return true;
  if (needle.length <= 3) return true;
  return false;
}

function addFindCandidate(candidates, item, context) {
  if (!shouldIncludeFindItem(item, context.args, context.extensionInfo)) return;
  const key = findCandidateKey(item);
  const scored = scoreFindCandidate(item, context);
  const queryTokens = findImportantTokens(context.query);
  const contentCoverage = queryTokens.length && context.contentMatchedTokens
    ? context.contentMatchedTokens / queryTokens.length
    : 0;
  const strongContentIndexMatch = context.source === "contentIndex"
    && (context.contentExactPhrase === true || (context.contentMatchedTokens || 0) >= 2 || contentCoverage >= 0.5);
  const canonicalGraphMatch = context.source === "search" && context.canonicalSearch === true;
  const hasQueryRelevance = scored.matchedTokens > 0 || scored.reasons.some((reason) =>
    reason.startsWith("exact filename")
    || reason.startsWith("filename contains")
    || reason.startsWith("path contains")
    || reason.startsWith("date match")
    || reason.startsWith("requested extension")
    || reason.startsWith("likely file type")
  ) || strongContentIndexMatch || canonicalGraphMatch;
  if (context.source !== "exactPath" && !hasQueryRelevance) return;
  const existing = candidates.get(key);
  if (!existing || scored.score > existing.score) {
    candidates.set(key, {
      item,
      score: scored.score,
      reasons: scored.reasons,
      snippets: context.snippet ? [context.snippet] : [],
      sources: [{ source: context.source, term: context.term, folder: context.folder }]
    });
  } else {
    existing.sources.push({ source: context.source, term: context.term, folder: context.folder });
    existing.reasons = [...new Set([...existing.reasons, ...scored.reasons])].slice(0, 6);
    if (context.snippet && !existing.snippets.includes(context.snippet)) existing.snippets.push(context.snippet);
  }
}

function rankedFindCandidates(candidates) {
  return [...candidates.values()].sort((left, right) => {
    if (right.score !== left.score) return right.score - left.score;
    return String(left.item.name || "").localeCompare(String(right.item.name || ""));
  });
}

function candidateHasLiveSource(candidate) {
  return candidate.sources.some((source) => ["exactPath", "search", "scan", "metadataConfirm"].includes(source.source));
}

function cacheConfirmationEligible(candidate, query = "") {
  const normalizedQuery = normalizeFindText(query);
  const nameText = normalizeFindText(candidate.item?.name || "");
  const pathText = normalizeFindText(candidate.item?.remotePath || candidate.item?.path || "");
  const queryLooksPathLike = String(query).includes("/");
  if (normalizedQuery && nameText === normalizedQuery) return true;
  if (queryLooksPathLike && normalizedQuery && pathText.includes(normalizedQuery)) return true;
  const extension = findItemExtension(candidate.item);
  return Boolean(extension && normalizedQuery.endsWith(extension) && nameText.includes(normalizedQuery));
}

async function confirmCachedFindCandidates(candidates, ranked, context = {}) {
  const maxConfirmations = clampInteger(context.args?.cacheConfirmMaxItems, 5, 1, 20);
  const minScore = clampInteger(context.args?.cacheConfirmMinScore, 60, 0, 200);
  const targets = ranked
    .filter((candidate) => !candidateHasLiveSource(candidate))
    .filter((candidate) => candidate.item?.id && candidate.score >= minScore)
    .filter((candidate) => candidate.sources.some((source) => source.source === "cache"))
    .filter((candidate) => cacheConfirmationEligible(candidate, context.query))
    .slice(0, maxConfirmations);
  if (!targets.length) return { attempted: 0, confirmed: 0, errors: 0 };
  let confirmed = 0;
  let errors = 0;
  let confirmedItems = [];
  try {
    const result = await batchGetInfo({
      items: targets.map((candidate) => ({ itemId: candidate.item.id })),
      format: "full"
    });
    for (const item of result.items || []) {
      if (item.error) {
        errors += 1;
        continue;
      }
      confirmedItems.push(item);
    }
  } catch {
    for (const candidate of targets) {
      try {
        confirmedItems.push(simplifyItem(await getRawInfo({ itemId: candidate.item.id })));
      } catch {
        errors += 1;
      }
    }
  }
  for (const item of confirmedItems) {
    if (item.error) {
      errors += 1;
      continue;
    }
    confirmed += 1;
    addFindCandidate(candidates, item, {
      args: context.args,
      query: context.query,
      source: "metadataConfirm",
      term: "batch-get-info",
      extensionInfo: context.extensionInfo,
      scoringFolderHints: context.scoringFolderHints
    });
  }
  return { attempted: targets.length, confirmed, errors };
}

function scanRootKey(root = {}, fallbackFolder = "") {
  return root.id || normalizeFolderHintKey(root.remotePath || fallbackFolder);
}

async function find(args = {}) {
  const startedAt = Date.now();
  const query = String(args.query || "").trim();
  if (!query) throw new Error("query is required.");

  const maxResults = Math.min(args.maxResults ?? 10, args.maxResultsLimit ?? 50);
  const maxSearchTerms = Math.min(args.maxSearchTerms ?? 8, 12);
  const searchTerms = buildFindSearchTerms(query, maxSearchTerms);
  const extensionInfo = inferFindExtensions(query, args.extensions || []);
  const scoringFolderHints = pruneFolderHints(args.folderHints || []);
  const folderHints = pruneFolderHints([...scoringFolderHints, ...defaultFindFolderHints(query), ""]);
  const candidates = new Map();
  const searchRuns = [];
  const scanRuns = [];
  const pendingSearchCacheItems = [];
  const searchConcurrency = clampInteger(args.searchConcurrency, 2, 1, 4);
  const searchConfidenceThreshold = args.minConfidenceForSearchOnly ?? 78;
  const executedSearchTerms = [];
  const skippedSearchTerms = [];
  let searchStopReason = null;
  let cacheCandidateCount = 0;
  let contentIndexCandidateCount = 0;
  let contentIndexDurationMs = 0;
  let liveSearchDurationMs = 0;
  let scanDurationMs = 0;
  let cacheConfirmDurationMs = 0;
  let cacheConfirmations = { attempted: 0, confirmed: 0, errors: 0 };

  if (args.useCache !== false) {
    const cacheList = await cachedItems();
    for (const item of cacheList) {
      addFindCandidate(candidates, item, {
        args,
        query,
        source: "cache",
        term: "metadata-cache",
        extensionInfo,
        scoringFolderHints
      });
    }
    cacheCandidateCount = cacheList.length;
  }

  if (args.useContentIndex !== false && (args.contentMaxResults ?? 10) > 0) {
    const contentStartedAt = Date.now();
    const indexed = await contentSearch({
      query,
      maxResults: clampInteger(args.contentMaxResults, 10, 0, 100) || 1,
      format: "full"
    });
    contentIndexDurationMs += elapsedMs(contentStartedAt);
    contentIndexCandidateCount = indexed.items?.length || 0;
    for (const match of indexed.items || []) {
      addFindCandidate(candidates, match, {
        args,
        query,
        source: "contentIndex",
        term: "content-index",
        extensionInfo,
        scoringFolderHints,
        contentScore: match.score,
        contentMatchedTokens: match.matchedTokens || 1,
        contentExactPhrase: match.reasons?.includes("indexed content phrase"),
        contentReason: match.reasons?.[0],
        snippet: match.snippet
      });
    }
  }

  if (query.includes("/")) {
    try {
      const item = simplifyItem(await getRawInfo({ path: query, cacheResults: args.useCache !== false }));
      addFindCandidate(candidates, item, {
        args,
        query,
        source: "exactPath",
        extensionInfo,
        scoringFolderHints
      });
    } catch (error) {
      searchRuns.push({ strategy: "exactPath", path: query, error: safeToolErrorMessage(error) });
    }
  }

  const executeSearchTerm = async ({ term, index, stage }) => {
    const termStartedAt = Date.now();
    executedSearchTerms.push(term);
    try {
      const result = await searchAll({
        query: term,
        pageSize: args.searchPageSize ?? 50,
        maxItems: args.searchMaxItemsPerTerm ?? 100,
        format: "full",
        cacheResults: false
      });
      if (args.useCache !== false) pendingSearchCacheItems.push(...(result.items || []));
      for (const [resultIndex, item] of (result.items || []).entries()) {
        addFindCandidate(candidates, item, {
          args,
          query,
          source: "search",
          term,
          canonicalSearch: index === 0,
          searchResultRank: resultIndex + 1,
          extensionInfo,
          scoringFolderHints
        });
      }
      return {
        term,
        stage,
        executed: true,
        count: result.count,
        truncated: result.truncated,
        unsafePageTruncation: result.unsafePageTruncation,
        graphSearchCalls: result.pagesFetched || 0,
        durationMs: elapsedMs(termStartedAt)
      };
    } catch (error) {
      return {
        term,
        stage,
        executed: true,
        graphSearchCalls: error.graphPagesFetched || 0,
        durationMs: elapsedMs(termStartedAt),
        error: safeToolErrorMessage(error)
      };
    }
  };

  const executeSearchWave = async (entries) => {
    if (!entries.length) return;
    const waveStartedAt = Date.now();
    const runs = await mapWithConcurrency(entries, searchConcurrency, executeSearchTerm);
    liveSearchDurationMs += elapsedMs(waveStartedAt);
    searchRuns.push(...runs);
  };

  if (searchTerms.length) {
    await executeSearchWave([{ term: searchTerms[0], index: 0, stage: "canonical" }]);
    let nextSearchTermIndex = 1;
    let bestLiveSearchScore = rankedFindCandidates(candidates).find(candidateHasLiveSource)?.score || 0;
    if (args.executeAllSearchTerms !== true && bestLiveSearchScore >= searchConfidenceThreshold && nextSearchTermIndex < searchTerms.length) {
      searchStopReason = "high-confidence-canonical";
    } else {
      while (nextSearchTermIndex < searchTerms.length) {
        const wave = searchTerms
          .slice(nextSearchTermIndex, nextSearchTermIndex + searchConcurrency)
          .map((term, offset) => ({ term, index: nextSearchTermIndex + offset, stage: "expansion" }));
        nextSearchTermIndex += wave.length;
        await executeSearchWave(wave);
        bestLiveSearchScore = rankedFindCandidates(candidates).find(candidateHasLiveSource)?.score || 0;
        if (args.executeAllSearchTerms !== true && bestLiveSearchScore >= searchConfidenceThreshold && nextSearchTermIndex < searchTerms.length) {
          searchStopReason = "high-confidence-expansion";
          break;
        }
      }
    }

    for (const term of searchTerms.slice(executedSearchTerms.length)) {
      skippedSearchTerms.push(term);
      searchRuns.push({
        term,
        stage: "expansion",
        executed: false,
        skipped: searchStopReason || "adaptive-stop"
      });
    }
  }
  if (!searchStopReason) searchStopReason = "all-terms-executed";

  if (args.useCache !== false && pendingSearchCacheItems.length) {
    const uniqueSearchCacheItems = [...new Map(pendingSearchCacheItems.filter((item) => item?.id).map((item) => [item.id, item])).values()];
    await bestEffortLocalWrite("metadata cache update", async () => await withMetadataCacheBatch(async () => {
      await cacheItems(uniqueSearchCacheItems);
    }));
  }

  let ranked = rankedFindCandidates(candidates);
  if (args.confirmCacheCandidates !== false) {
    const confirmStartedAt = Date.now();
    cacheConfirmations = await confirmCachedFindCandidates(candidates, ranked, {
      args,
      query,
      extensionInfo,
      scoringFolderHints
    });
    cacheConfirmDurationMs += elapsedMs(confirmStartedAt);
    ranked = rankedFindCandidates(candidates);
  }
  const bestLiveScore = ranked.find(candidateHasLiveSource)?.score || 0;
  const shouldScan = args.scanFallback !== false && bestLiveScore < searchConfidenceThreshold;

  if (shouldScan) {
    const scanNeedle = findScanNeedle(query);
    const scanConcurrency = clampInteger(args.scanConcurrency, 2, 1, 4);
    const scanPlans = [
      { nameContains: scanNeedle, extensions: [...extensionInfo.effectiveForScan], reason: "targeted" }
    ];
    if (extensionInfo.effectiveForScan.size || scanNeedleIsWeak(scanNeedle, query)) {
      scanPlans.push({ nameContains: scanNeedle, extensions: [], reason: "no-extension-filter" });
    }
    if (scanNeedle && scanNeedleIsWeak(scanNeedle, query) && extensionInfo.inferred.size) {
      scanPlans.push({ nameContains: "", extensions: [...extensionInfo.strictInferred], reason: "type-only" });
    }

    for (const plan of scanPlans) {
      let remainingItems = Math.min(args.scanMaxItems ?? 1500, 10000);
      let remainingFolders = Math.min(args.scanMaxFolders ?? 250, 2000);
      const scannedFolderKeys = new Set();
      const completedScanRootIds = new Set();
      let folderIndex = 0;
      while (folderIndex < folderHints.length && remainingItems > 0 && remainingFolders > 0) {
        const targetBatch = [];
        while (folderIndex < folderHints.length && targetBatch.length < scanConcurrency) {
          const folder = folderHints[folderIndex];
          folderIndex += 1;
          // Keep the broad root fallback in its own wave so it can exclude every
          // hinted subtree that earlier waves already traversed completely.
          if (!folder && targetBatch.length > 0) {
            folderIndex -= 1;
            break;
          }
          try {
            const root = await resolveScanRoot({
              ...(folder ? { path: folder } : {}),
              useCache: args.useCache !== false,
              cacheResults: args.useCache !== false
            });
            const key = scanRootKey(root, folder);
            if (scannedFolderKeys.has(key)) {
              scanRuns.push({ folder: folder || "root", reason: plan.reason, skipped: "duplicate-root", root: { id: root.id, remotePath: root.remotePath } });
              continue;
            }
            scannedFolderKeys.add(key);
            targetBatch.push({ folder, root });
          } catch (error) {
            scanRuns.push({ folder: folder || "root", reason: plan.reason, error: safeToolErrorMessage(error) });
          }
        }
        if (!targetBatch.length) continue;
        const perScanMaxItems = Math.max(1, Math.ceil(remainingItems / targetBatch.length));
        const perScanMaxFolders = Math.max(1, Math.ceil(remainingFolders / targetBatch.length));
        const scanStartedAt = Date.now();
        const batchResults = await mapWithConcurrency(targetBatch, scanConcurrency, async (target) => {
          try {
            const result = await scan({
              path: target.folder,
              _resolvedRoot: target.root,
              nameContains: plan.nameContains,
              extensions: plan.extensions,
              includeFiles: true,
              includeFolders: args.includeFolders === true,
              maxItems: perScanMaxItems,
              maxFolders: perScanMaxFolders,
              maxDepth: args.scanMaxDepth ?? 20,
              maxResults: Math.max(maxResults * 2, 20),
              stopAfterResults: true,
              format: "full",
              cacheResults: args.useCache !== false,
              _skipFolderIds: [...completedScanRootIds].filter((id) => id !== target.root.id)
            });
            return { target, result };
          } catch (error) {
            return { target, error };
          }
        });
        scanDurationMs += elapsedMs(scanStartedAt);
        for (const entry of batchResults) {
          const folderLabel = entry.target.folder || "root";
          if (entry.error) {
            scanRuns.push({ folder: folderLabel, reason: plan.reason, error: safeToolErrorMessage(entry.error) });
            continue;
          }
          const result = entry.result;
          scanRuns.push({ folder: folderLabel, reason: plan.reason, nameContains: plan.nameContains || null, extensions: plan.extensions, summary: result.summary });
          if (!result.summary.traversalTruncated && entry.target.root.id) {
            completedScanRootIds.add(entry.target.root.id);
          }
          remainingItems -= result.summary.itemsScanned || 0;
          remainingFolders -= result.summary.foldersVisited || 0;
          for (const item of result.items || []) {
            addFindCandidate(candidates, item, {
              args,
              query,
              source: "scan",
              folder: folderLabel,
              extensionInfo,
              scoringFolderHints
            });
          }
        }
        ranked = rankedFindCandidates(candidates);
        if ((ranked[0]?.score || 0) >= searchConfidenceThreshold && ranked.length >= maxResults) break;
      }
      ranked = rankedFindCandidates(candidates);
      if (ranked.length > 0) break;
    }
  }

  ranked = rankedFindCandidates(candidates);
  if (shouldScan) {
    ranked = ranked.filter(candidateHasLiveSource);
  }
  const noteParts = [];
  if (cacheCandidateCount > 0) noteParts.push("Used the local metadata cache as a candidate source.");
  if (contentIndexCandidateCount > 0) noteParts.push("Used matching entries from the local content index.");
  if (executedSearchTerms.length > 0) {
    noteParts.push(`Ran ${executedSearchTerms.length} live Graph search ${executedSearchTerms.length === 1 ? "term" : "terms"}.`);
  }
  if (skippedSearchTerms.length > 0) {
    noteParts.push(`Skipped ${skippedSearchTerms.length} expansion ${skippedSearchTerms.length === 1 ? "term" : "terms"} after reaching the confidence threshold.`);
  }
  if (shouldScan) {
    noteParts.push("Used bounded remote recursive scan fallback because live-search confidence remained below the threshold.");
  } else if (args.scanFallback === false) {
    noteParts.push("Recursive scan fallback was disabled.");
  } else {
    noteParts.push("Recursive scan fallback was not required by the configured confidence threshold.");
  }
  return {
    query,
    strategy: "cache-assisted-remote-first",
    searchTerms,
    searchPlan: {
      mode: "adaptive-staged",
      concurrency: searchConcurrency,
      confidenceThreshold: searchConfidenceThreshold,
      planned: searchTerms,
      executed: executedSearchTerms,
      skipped: skippedSearchTerms,
      stopReason: searchStopReason
    },
    inferred: {
      extensions: [...extensionInfo.inferred],
      strictExtensions: [...extensionInfo.strictInferred],
      matchedKinds: extensionInfo.matchedKinds,
      explicitExtensions: [...extensionInfo.explicit]
    },
    summary: {
      candidates: ranked.length,
      returned: Math.min(ranked.length, maxResults),
      bestScore: ranked[0]?.score || 0,
      usedScanFallback: shouldScan,
      scanRuns: scanRuns.length,
      localIndexUsed: contentIndexCandidateCount > 0,
      contentIndexCandidates: contentIndexCandidateCount,
      persistentCacheUsed: cacheCandidateCount > 0,
      cacheCandidates: cacheCandidateCount,
      durationMs: elapsedMs(startedAt),
      contentIndexDurationMs,
      liveSearchDurationMs,
      scanDurationMs,
      cacheConfirmDurationMs,
      cacheConfirmations,
      graphSearchCalls: searchRuns.reduce((sum, run) => sum + (run.graphSearchCalls || 0), 0),
      searchTermsPlanned: searchTerms.length,
      searchTermsExecuted: executedSearchTerms.length,
      searchTermsSkipped: skippedSearchTerms.length,
      searchStopReason,
      scanAttempts: scanRuns.length,
      metadataCacheWrites: toolCallContext.getStore()?.metadataCacheWrites || 0
    },
    searchRuns,
    scanRuns,
    items: ranked.slice(0, maxResults).map((candidate) => ({
      ...formatSimplifiedItem(candidate.item, args.format),
      score: candidate.score,
      reasons: candidate.reasons,
      snippets: candidate.snippets?.slice(0, 2),
      sources: candidate.sources.slice(0, 5)
    })),
    note: noteParts.join(" ")
  };
}

function broadFindFolderHints(query = "", userHints = []) {
  const presets = pathPresets();
  const common = [
    presets.documents,
    presets.desktop,
    presets.pictures,
    presets.screenshots,
    "Personal/Documents",
    "Personal",
    "Microsoft Copilot Chat Files",
    ""
  ];
  return pruneFolderHints([
    ...(userHints || []),
    ...defaultFindFolderHints(query),
    ...common
  ].filter((hint) => hint !== undefined && hint !== null));
}

async function findAll(args = {}) {
  const query = String(args.query || "").trim();
  if (!query) throw new Error("query is required.");
  const result = await find({
    ...args,
    maxResults: Math.min(args.maxResults ?? 50, 200),
    maxResultsLimit: 200,
    folderHints: broadFindFolderHints(query, args.folderHints || []),
    scanFallback: true,
    scanMaxItems: Math.min(args.scanMaxItems ?? 10000, 50000),
    scanMaxFolders: Math.min(args.scanMaxFolders ?? 2000, 10000),
    scanMaxDepth: Math.min(args.scanMaxDepth ?? 25, 50),
    searchPageSize: Math.min(args.searchPageSize ?? 100, 200),
    searchMaxItemsPerTerm: Math.min(args.searchMaxItemsPerTerm ?? 250, 1000),
    minConfidenceForSearchOnly: args.minConfidenceForSearchOnly ?? 78,
    executeAllSearchTerms: true
  });
  return {
    ...result,
    strategy: "broad-cache-assisted-remote-first",
    folderPlan: broadFindFolderHints(query, args.folderHints || []).map((folder) => folder || "root")
  };
}

function formatDeltaItem(item, format = "compact") {
  const formatted = formatDriveItem(item, format);
  if (format === "full") return formatted ? { ...formatted, deleted: item.deleted } : formatted;
  return formatted ? { ...formatted, deleted: item.deleted ? item.deleted : undefined } : formatted;
}

async function resolveDeltaPathRoot(args = {}) {
  try {
    if (args.itemId) return await resolveScanRoot({ itemId: args.itemId, cacheResults: false });
    const resolvedPath = resolvePresetPath(args);
    if (resolvedPath) return await resolveScanRoot({ path: resolvedPath, cacheResults: false });
    return await resolveScanRoot({ cacheResults: false });
  } catch (error) {
    recordLocalWarning("delta path-root resolution", error);
    return null;
  }
}

async function hydrateDeltaParentItems(items = [], pathRoot = null) {
  if (!items.length) return items;
  const cache = await loadMetadataCache();
  const combined = new Map(items.filter((item) => item?.id).map((item) => [item.id, item]));
  const supplemental = new Map();
  const knownRootIds = new Set([
    ...Object.keys(cache.pathRootsById || {}),
    ...(cache.scanRoot?.id ? [cache.scanRoot.id] : []),
    ...(pathRoot?.id ? [pathRoot.id] : [])
  ]);
  const maxHydratedParents = Math.min(100, Math.max(10, items.length * 2));

  while (supplemental.size < maxHydratedParents) {
    const missingParentIds = [];
    for (const item of combined.values()) {
      if (!item?.parentReference?.id || itemRemotePath(item) !== undefined) continue;
      const parentId = item.parentReference.id;
      if (knownRootIds.has(parentId)) continue;
      const cachedParent = cache.itemsById?.[parentId];
      if (cachedParent?.remotePath !== undefined) continue;
      if (combined.has(parentId)) continue;
      if (!missingParentIds.includes(parentId)) missingParentIds.push(parentId);
      if (supplemental.size + missingParentIds.length >= maxHydratedParents) break;
    }
    if (!missingParentIds.length) break;
    const fetched = await mapWithConcurrency(missingParentIds, 4, async (parentId) => {
      try {
        return await graph(`/me/drive/items/${encodeURIComponent(parentId)}?$select=${encodeURIComponent(defaultSelect)}`);
      } catch (error) {
        recordLocalWarning("delta parent path resolution", error);
        return null;
      }
    });
    let added = 0;
    for (const parent of fetched) {
      if (!parent?.id || combined.has(parent.id)) continue;
      combined.set(parent.id, parent);
      supplemental.set(parent.id, parent);
      added += 1;
    }
    if (!added) break;
  }
  return [...supplemental.values(), ...items];
}

async function delta(args = {}) {
  const maxItems = clampInteger(args.maxItems, 1000, 1, 5000);
  let firstPath = args.nextLink || args.deltaLink;
  let target = args._deltaTarget || (args.nextLink ? "nextLink" : args.deltaLink ? "deltaLink" : "root");
  let pathRoot = null;
  if (firstPath && !isDeltaCursor(firstPath)) {
    throw new Error("nextLink and deltaLink must be Microsoft Graph delta cursor URLs. Refusing a non-delta pagination cursor.");
  }
  if (!firstPath) {
    const params = new URLSearchParams();
    params.set("$top", String(Math.min(clampInteger(args.pageSize, 200, 1, 200), maxItems)));
    if (args.itemId) {
      firstPath = `/me/drive/items/${encodeURIComponent(args.itemId)}/delta?${params.toString()}`;
      target = `itemId:${args.itemId}`;
    } else {
      const resolvedPath = resolvePresetPath(args);
      if (resolvedPath) {
        pathRoot = await resolveDeltaPathRoot(args);
        if (!pathRoot?.id) throw new Error(`Could not resolve delta target folder: ${resolvedPath}`);
        firstPath = `/me/drive/items/${encodeURIComponent(pathRoot.id)}/delta?${params.toString()}`;
        target = resolvedPath;
      } else {
        firstPath = `/me/drive/root/delta?${params.toString()}`;
      }
    }
    if (!pathRoot) pathRoot = await resolveDeltaPathRoot(args);
    if (pathRoot) {
      await bestEffortLocalWrite("metadata cache path-root update", async () => await cacheItems([], { pathRoot }));
    }
  }
  const result = await collectPages(firstPath, maxItems, args.format, formatDeltaItem, {
    cacheResults: true,
    persistDeltaCursor: args._persistCursor === true,
    deltaTarget: target,
    prepareCacheItems: async (items) => await hydrateDeltaParentItems(items, pathRoot)
  });
  const cache = await loadMetadataCache();
  const unresolvedPathCount = unresolvedPathItems(cache).length;
  const resolvedItems = result.items.map((item) => {
    if (!item?.id || item.deleted) return item;
    const cached = cache.itemsById?.[item.id];
    return cached ? formatSimplifiedItem(cached, args.format) : item;
  });
  return {
    ...result,
    items: resolvedItems,
    target,
    unresolvedPathCount,
    note: result.unsafePageTruncation
      ? "Microsoft Graph returned more items than could be accepted safely from one page. No continuation was returned because it would skip items; rerun with a larger maxItems."
      : result.deltaLink
        ? "Save deltaLink to ask for changes since this point later."
        : "Use nextLink to continue this delta scan before saving a deltaLink."
  };
}

async function cacheRefresh(args = {}) {
  const startedAt = Date.now();
  const existing = await loadMetadataCache();
  const invalidCursorMetadata = {};
  if (existing.deltaLink && !isDeltaCursor(existing.deltaLink)) {
    existing.deltaLink = null;
    invalidCursorMetadata.deltaLink = null;
  }
  if (existing.deltaNextLink && !isDeltaCursor(existing.deltaNextLink)) {
    existing.deltaNextLink = null;
    invalidCursorMetadata.deltaNextLink = null;
  }
  if ((!existing.deltaLink && !existing.deltaNextLink) && existing.deltaTarget) {
    existing.deltaTarget = null;
    invalidCursorMetadata.deltaTarget = null;
  }
  if (Object.keys(invalidCursorMetadata).length) {
    await cacheItems([], invalidCursorMetadata);
  }
  const settings = pluginSettings();
  const mode = args.mode || "auto";
  const progress = [];
  const addProgress = (stage, details = {}) => {
    progress.push({ stage, elapsedMs: elapsedMs(startedAt), ...details });
  };
  const requestedTarget = args.itemId
    ? `itemId:${args.itemId}`
    : (resolvePresetPath(args) || "root");
  const cachedTarget = existing.deltaTarget || existing.scanRoot?.target || "root";
  const targetMatchesCache = requestedTarget === cachedTarget
    && (!existing.deltaTarget || existing.deltaTarget === requestedTarget);
  if ((mode === "delta" || mode === "auto") && settings.deltaSyncEnabled && (existing.deltaNextLink || existing.deltaLink) && targetMatchesCache) {
    const continuingNextLink = Boolean(existing.deltaNextLink);
    addProgress(continuingNextLink ? "delta-resume-nextLink" : "delta-start", {
      cursor: continuingNextLink ? "deltaNextLink" : "deltaLink"
    });
    const result = await delta({
      ...(continuingNextLink ? { nextLink: existing.deltaNextLink } : { deltaLink: existing.deltaLink }),
      pageSize: clampInteger(args.pageSize, 200, 1, 200),
      maxItems: clampInteger(args.maxItems, 10000, 1, 50000),
      format: "full",
      _persistCursor: true,
      _deltaTarget: requestedTarget
    });
    addProgress(result.deltaLink ? "delta-complete" : "delta-incomplete", {
      count: result.count,
      hasNextLink: Boolean(result.nextLink),
      hasDeltaLink: Boolean(result.deltaLink)
    });
    return {
      mode: "delta",
      result,
      cache: await syncStatus(),
      progress,
      durationMs: elapsedMs(startedAt),
      note: result.deltaLink ? "Applied delta changes and stored the latest deltaLink." : "Delta scan is incomplete; run again with nextLink before treating cache as fresh."
    };
  }

  if (mode === "delta") {
    if (!settings.deltaSyncEnabled) {
      throw new Error("Delta sync is disabled by configuration.");
    }
    if (!existing.deltaLink && !existing.deltaNextLink) {
      throw new Error("No cached deltaLink or deltaNextLink exists yet. Run onedrive_cache_refresh with mode: scan or auto first.");
    }
    throw new Error(`Cached deltaLink is for ${cachedTarget}, not requested target ${requestedTarget}. Run onedrive_cache_refresh with mode: scan for this target.`);
  }

  if (args.replaceCache === true) await clearMetadataCache();
  addProgress("scan-start", { target: requestedTarget, replaceCache: args.replaceCache === true });
  const result = await scan({
    ...args,
    format: "full",
    includeFiles: true,
    includeFolders: args.includeFolders !== false,
    maxItems: clampInteger(args.maxItems, 10000, 1, 50000),
    maxFolders: clampInteger(args.maxFolders, 2000, 1, 10000),
    maxDepth: clampInteger(args.maxDepth, settings.maxScanDepth, 0, 50)
  });
  addProgress("scan-complete", {
    itemsScanned: result.summary?.itemsScanned,
    foldersVisited: result.summary?.foldersVisited,
    traversalTruncated: result.summary?.traversalTruncated
  });
  await cacheItems([], {
    scanRoot: result.root,
    deltaLink: null,
    deltaNextLink: null,
    deltaTarget: requestedTarget
  });

  try {
    if (settings.deltaSyncEnabled) {
      addProgress("delta-prime-start");
      const deltaResult = await delta({
        ...args,
        pageSize: clampInteger(args.pageSize, 200, 1, 200),
        maxItems: clampInteger(args.maxItems, 10000, 1, 50000),
        format: "full",
        _persistCursor: true,
        _deltaTarget: requestedTarget
      });
      const deltaComplete = Boolean(deltaResult.deltaLink);
      addProgress(deltaComplete ? "delta-prime-complete" : "delta-prime-incomplete", {
        count: deltaResult.count,
        hasNextLink: Boolean(deltaResult.nextLink),
        hasDeltaLink: Boolean(deltaResult.deltaLink)
      });
      return {
        mode,
        effectiveMode: "scan",
        scan: result,
        deltaRefresh: { attempted: true, complete: deltaComplete, count: deltaResult.count, deltaLink: deltaResult.deltaLink, nextLink: deltaResult.nextLink },
        cache: await syncStatus(),
        progress,
        durationMs: elapsedMs(startedAt),
        note: deltaComplete
          ? `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan and stored a fresh deltaLink.`
          : `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan. Delta refresh returned nextLink but no final deltaLink yet.`
      };
    }
    return {
      mode: "scan",
      effectiveMode: "scan",
      scan: result,
      cache: await syncStatus(),
      progress,
      durationMs: elapsedMs(startedAt),
      note: `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan. Delta sync is disabled by configuration.`
    };
  } catch (error) {
    addProgress("delta-prime-error", { error: safeToolErrorMessage(error) });
    return {
      mode: "scan",
      effectiveMode: "scan",
      scan: result,
      cache: await syncStatus(),
      deltaRefresh: { attempted: true, complete: false },
      deltaError: safeToolErrorMessage(error),
      progress,
      durationMs: elapsedMs(startedAt),
      note: `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan, but could not get a fresh deltaLink.`
    };
  }
}

function batchResponseHeader(headers = {}, name) {
  if (typeof headers?.get === "function") return headers.get(name);
  const target = String(name).toLowerCase();
  const entry = Object.entries(headers || {}).find(([key]) => key.toLowerCase() === target);
  return entry?.[1] ?? null;
}

function batchRetryDelayMs(responses, attempt) {
  const delays = responses.map((response) => {
    const retryAfter = batchResponseHeader(response.headers, "retry-after");
    if (retryAfter === null || retryAfter === undefined || retryAfter === "") return null;
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const dateMs = Date.parse(retryAfter);
    return Number.isFinite(dateMs) ? Math.max(0, dateMs - Date.now()) : null;
  }).filter((value) => value !== null);
  if (delays.length) return Math.max(...delays);
  return Math.min(1000 * 2 ** attempt, 8000);
}

function isReadOnlyBatchRequest(request = {}) {
  return ["GET", "HEAD"].includes(String(request.method || "GET").toUpperCase());
}

function isTransientBatchResponse(response = {}) {
  return response.status === 0 || [429, 500, 502, 503, 504].includes(response.status);
}

async function batchGraph(requests = []) {
  if (!requests.length) return [];
  if (requests.length > 20) throw new Error("Microsoft Graph batch requests support at most 20 subrequests.");
  const entries = requests.map((request, index) => ({ request, index, id: String(index + 1) }));
  const finalResponses = new Array(requests.length);
  let pending = entries;
  const maxRetries = 3;

  for (let attempt = 0; pending.length; attempt += 1) {
    const allReadOnly = pending.every((entry) => isReadOnlyBatchRequest(entry.request));
    const result = await graph("/$batch", {
      method: "POST",
      isMutation: !allReadOnly,
      body: JSON.stringify({
        requests: pending.map(({ request, id }) => ({
          id,
          method: request.method || "GET",
          url: String(request.url || "").replace(/^\/+/, ""),
          ...(request.headers ? { headers: request.headers } : {}),
          ...(request.body !== undefined ? { body: request.body } : {})
        }))
      }),
      ...(allReadOnly ? { maxRetries: 3 } : {})
    });
    const responses = new Map((result.responses || []).map((response) => [String(response.id), response]));
    const retryEntries = [];
    const retryResponses = [];

    for (const entry of pending) {
      const response = responses.get(entry.id);
      const normalized = response ? {
        ok: response.status >= 200 && response.status < 300,
        status: response.status,
        body: response.body,
        headers: response.headers,
        request: entry.request
      } : {
        ok: false,
        status: 0,
        error: "Missing batch response.",
        request: entry.request
      };
      if (attempt < maxRetries && isReadOnlyBatchRequest(entry.request) && isTransientBatchResponse(normalized)) {
        retryEntries.push(entry);
        retryResponses.push(normalized);
      } else {
        finalResponses[entry.index] = normalized;
      }
    }

    if (!retryEntries.length) break;
    await sleep(batchRetryDelayMs(retryResponses, attempt));
    pending = retryEntries;
  }

  return finalResponses;
}

async function batchGetInfo(args = {}) {
  const items = args.items || [];
  const responses = await batchGraph(items.map((target) => ({ url: itemBase(target), target })));
  const rawItems = responses.filter((response) => response.ok && response.body).map((response) => response.body);
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(rawItems));
  return {
    count: responses.length,
    items: responses.map((response) => response.ok
      ? formatDriveItem(response.body, args.format || "compact")
      : { error: redactAuditText(response.body?.error?.message || response.error || `HTTP ${response.status}`), status: response.status, target: response.request.target })
  };
}

async function batchPermissions(args = {}) {
  const items = args.items || [];
  const responses = await batchGraph(items.map((target) => ({ url: `${itemBase(target)}/permissions`, target })));
  return {
    count: responses.length,
    items: responses.map((response) => response.ok
      ? {
          target: response.request.target,
          permissions: (response.body?.value || []).map((permission) => simplifyPermission(permission, args.format)),
          count: (response.body?.value || []).length
        }
      : { error: redactAuditText(response.body?.error?.message || response.error || `HTTP ${response.status}`), status: response.status, target: response.request.target })
  };
}

function uniqueBatchLocalPath(destinationFolder, name, index, usedTargets) {
  const parsed = parse(name || `onedrive-download-${index + 1}`);
  const base = parsed.name || "onedrive-download";
  const ext = parsed.ext || "";
  let suffix = 1;
  let candidate;
  do {
    const fileName = suffix === 1 ? `${base}${ext}` : `${base} (${suffix})${ext}`;
    candidate = resolve(join(destinationFolder, fileName));
    suffix += 1;
  } while (usedTargets.has(candidate));
  usedTargets.add(candidate);
  return candidate;
}

async function batchDownload(args = {}) {
  const destinationFolder = args.destinationFolder ? resolve(args.destinationFolder) : null;
  if (destinationFolder) await assertNotLocalOneDriveSyncPathForWrite(destinationFolder, "Batch download", args);
  const generatedDestinationFolder = destinationFolder || downloadRoot;
  const plannedTargets = new Set();
  const results = [];
  for (const [index, item] of (args.items || []).entries()) {
    try {
      const explicitLocalPath = item.localPath ? resolve(item.localPath) : null;
      if (explicitLocalPath) await assertNotLocalOneDriveSyncPathForWrite(explicitLocalPath, "Batch download", args);
      const info = await getInfo(item);
      const localPath = explicitLocalPath
        ? explicitLocalPath
        : uniqueBatchLocalPath(generatedDestinationFolder, info.name, index, plannedTargets);
      if (localPath) {
        await assertNotLocalOneDriveSyncPathForWrite(localPath, "Batch download", args);
        if (explicitLocalPath) {
          if (plannedTargets.has(localPath)) throw new Error(`Batch download target is used more than once: ${localPath}`);
          plannedTargets.add(localPath);
        }
      }
      results.push(await downloadResolvedItem(info, {
        itemId: info.id,
        localPath,
        overwrite: item.overwrite ?? args.overwrite,
        allowLocalOneDriveSyncPath: args.allowLocalOneDriveSyncPath
      }));
    } catch (error) {
      results.push({ target: item, error: safeToolErrorMessage(error) });
    }
  }
  return { count: results.length, results };
}

async function batchDelete(args = {}) {
  const items = args.items || [];
  const warnings = batchMutationWarnings();
  if (args.dryRun === false) {
    if (args.confirmed !== true) {
      return {
        dryRun: false,
        confirmed: false,
        count: items.length,
        warnings,
        requiredToDelete: "Set dryRun: false and confirmed: true after explicit user confirmation."
      };
    }
    const missingExpected = items.filter((item) => !item.expectedName && !item.expectedId);
    if (missingExpected.length) {
      return {
        dryRun: false,
        confirmed: true,
        count: items.length,
        warnings,
        requiredToDelete: "Provide expectedName or expectedId for every item in a live batch delete.",
        missingExpectedCount: missingExpected.length
      };
    }
  }

  const preflight = [];
  const preflightErrors = [];
  for (const [index, item] of items.entries()) {
    try {
      requireNonRootTarget(item, "Delete");
      const rawItem = await getRawInfo(item);
      if (rawItem.root) throw new Error("Delete refuses to operate on the OneDrive root.");
      assertExpectedItem(rawItem, item, "Delete");
      preflight.push({ targetArgs: item, rawItem, item: simplifyItem(rawItem) });
    } catch (error) {
      preflightErrors.push({ index, target: item, error: safeToolErrorMessage(error) });
    }
  }
  if (preflightErrors.length) {
    return {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: preflightErrors,
      requiredToDelete: "Fix every preflight error before running a batch delete."
    };
  }
  const previewProof = {
    items: preflight.map((entry) => ({ id: entry.rawItem.id, name: entry.rawItem.name })),
    operation: "batch-delete"
  };

  if (args.dryRun !== false) {
    return {
      dryRun: true,
      confirmed: args.confirmed === true,
      count: preflight.length,
      warnings,
      ...issuePreviewToken("onedrive_batch_delete", previewProof),
      results: preflight.map((entry) => ({ wouldDelete: entry.item }))
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(
    { dryRun: false, confirmed: true, count: preflight.length, warnings, results: preflight.map((entry) => ({ wouldDelete: entry.item })) },
    "onedrive_batch_delete",
    previewProof,
    args.previewToken,
    "requiredToDelete"
  );
  if (previewTokenRequired) return previewTokenRequired;

  const results = [];
  for (const [index, entry] of preflight.entries()) {
    try {
      await graph(itemMutationBase(entry.rawItem), { method: "DELETE", headers: mutationMatchHeaders(entry.rawItem) });
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([{ ...entry.rawItem, deleted: {} }]));
      results.push({ deleted: entry.item });
    } catch (error) {
      const failure = {
        dryRun: false,
        confirmed: true,
        count: preflight.length,
        warnings,
        failed: true,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        partialResults: results
      };
      await writeMutationAudit("onedrive_batch_delete", {
        status: "failed",
        targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
        partialResults: results.map((result) => ({ deleted: itemAuditSummary(result.deleted) })),
        failedIndex: index,
        error: safeErrorInfo(error)
      });
      return failure;
    }
  }
  await writeMutationAudit("onedrive_batch_delete", {
    status: "success",
    targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
    results: results.map((result) => ({ deleted: itemAuditSummary(result.deleted) }))
  });
  return { dryRun: false, confirmed: true, count: results.length, warnings, results };
}

async function getRawInfo(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (!resolved.path && !resolved.itemId) throw new Error("Provide path, preset, or itemId.");
  if (!args.includeDeletedItems && args.useCache === true) {
    const cached = resolved.itemId ? await cachedItemById(resolved.itemId) : await cachedItemByPath(resolved.path);
    if (cached) return cached;
  }
  const suffix = args.includeDeletedItems && resolved.itemId ? "?includeDeletedItems=true" : "";
  const item = await graph(`${itemBase(args)}${suffix}`);
  if (args.cacheResults !== false) {
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([item]));
  }
  return item;
}

async function getInfo(args = {}) {
  return formatDriveItem(await getRawInfo(args), args.format || "full");
}

function compactIdentity(identity) {
  if (!identity) return undefined;
  const typed = identity.user ? { type: "user", ...identity.user }
    : identity.group ? { type: "group", ...identity.group }
      : identity.siteUser ? { type: "siteUser", ...identity.siteUser }
        : identity.siteGroup ? { type: "siteGroup", ...identity.siteGroup }
          : identity.application ? { type: "application", ...identity.application }
            : identity.device ? { type: "device", ...identity.device }
              : null;
  return typed ? {
    type: typed.type,
    id: typed.id,
    displayName: typed.displayName,
    email: typed.email,
    loginName: typed.loginName
  } : identity;
}

function simplifyPermission(permission, format = "compact") {
  if (format === "full") return permission;
  const roles = new Set(permission.roles || []);
  const permissionKind = roles.has("owner")
    ? "owner"
    : permission.inheritedFrom
      ? "inherited"
      : permission.link?.scope === "anonymous"
        ? "anonymous_link"
        : permission.link
          ? "sharing_link"
          : permission.invitation
            ? "invitation"
            : "direct";
  const revocable = permissionKind !== "owner" && permissionKind !== "inherited";
  const link = permission.link && (permission.link.type || permission.link.scope) ? {
    type: permission.link.type,
    scope: permission.link.scope,
    webUrl: permission.link.webUrl,
    preventsDownload: permission.link.preventsDownload
  } : undefined;
  return {
    id: permission.id,
    roles: permission.roles,
    permissionKind,
    revocable,
    link,
    grantedTo: compactIdentity(permission.grantedTo),
    grantedToV2: compactIdentity(permission.grantedToV2),
    grantedToIdentities: permission.grantedToIdentities?.map(compactIdentity),
    grantedToIdentitiesV2: permission.grantedToIdentitiesV2?.map(compactIdentity),
    invitation: permission.invitation ? {
      email: permission.invitation.email,
      signInRequired: permission.invitation.signInRequired
    } : undefined,
    inheritedFrom: permission.inheritedFrom ? simplifyItem(permission.inheritedFrom) : undefined,
    expirationDateTime: permission.expirationDateTime,
    hasPassword: permission.hasPassword
  };
}

async function permissions(args = {}) {
  const result = await graph(`${itemBase(args)}/permissions`);
  return {
    item: await getInfo({ ...args, format: args.format || "compact" }).catch(() => null),
    permissions: (result.value || []).map((permission) => simplifyPermission(permission, args.format)),
    count: (result.value || []).length
  };
}

async function permissionList(args = {}, format = "compact") {
  const result = await graph(`${itemBase(args)}/permissions`);
  return (result.value || []).map((permission) => simplifyPermission(permission, format));
}

function permissionKey(permission = {}) {
  return permission.id
    || permission.link?.webUrl
    || `${(permission.roles || []).join(",")}:${permission.link?.type || ""}:${permission.link?.scope || ""}:${permission.grantedTo?.email || permission.invitation?.email || ""}`;
}

function isExplicitSharingPermission(permission = {}) {
  if (permission.inheritedFrom) return false;
  if (permission.permissionKind === "owner") return false;
  const roles = new Set(permission.roles || []);
  if (roles.has("owner")) return false;
  if (permission.link?.type || permission.link?.scope || permission.invitation) return true;
  return Boolean(
    permission.grantedTo
    || permission.grantedToV2
    || permission.grantedToIdentities?.length
    || permission.grantedToIdentitiesV2?.length
  );
}

function isOwnerPermission(permission = {}) {
  return permission.permissionKind === "owner" || new Set(permission.roles || []).has("owner");
}

function isRevocablePermission(permission = {}) {
  return permission.revocable !== false && !isOwnerPermission(permission) && !permission.inheritedFrom;
}

function diffPermissions(before = [], after = []) {
  const beforeMap = new Map(before.map((permission) => [permissionKey(permission), permission]));
  const afterMap = new Map(after.map((permission) => [permissionKey(permission), permission]));
  const added = [];
  const removed = [];
  const unchanged = [];
  for (const [key, permission] of afterMap) {
    if (beforeMap.has(key)) unchanged.push(permission);
    else added.push(permission);
  }
  for (const [key, permission] of beforeMap) {
    if (!afterMap.has(key)) removed.push(permission);
  }
  return {
    added,
    removed,
    unchangedCount: unchanged.length,
    beforeCount: before.length,
    afterCount: after.length
  };
}

function isTextMimeType(mimeType = "") {
  const lower = mimeType.toLowerCase();
  return textMimeTypes.has(lower) || textMimePrefixes.some((prefix) => lower.startsWith(prefix));
}

function isLikelyTextItem(info, args = {}) {
  const mimeType = info.file?.mimeType || "";
  if (mimeType && isTextMimeType(mimeType)) return true;
  const name = info.name || pathName(args.path || "");
  return textExtensions.has(extname(name).toLowerCase());
}

function assertTextReadable(info, args = {}) {
  if (info.folder) throw new Error(`Item is a folder, not a text file: ${info.name}`);
  if (args.force === true) return;
  if (!isLikelyTextItem(info, args)) {
    const mime = info.file?.mimeType || "unknown MIME type";
    throw new Error(`Refusing to read likely binary file as text (${mime}). Use onedrive_download, or pass force: true with maxBytes if you are sure.`);
  }
}

function assertNoBinaryNulls(buffer) {
  const sampleLength = Math.min(buffer.length, 4096);
  for (let index = 0; index < sampleLength; index += 1) {
    if (buffer[index] === 0) {
      throw new Error("Downloaded content contains NUL bytes and looks binary. Use onedrive_download instead.");
    }
  }
}

async function resolveDestinationParent(args = {}) {
  assertAtMostOneSelector(args, "destination parent", [
    { label: "destinationParentItemId", keys: ["destinationParentItemId"] },
    { label: "destinationParentPath", keys: ["destinationParentPath"] },
    { label: "destinationParentPreset", keys: ["destinationParentPreset"] },
    { label: "parentItemId", keys: ["parentItemId"] },
    { label: "parentPath", keys: ["parentPath"] },
    { label: "parentPreset", keys: ["parentPreset"] }
  ]);
  if (args.destinationParentItemId) return { id: args.destinationParentItemId };
  if (args.parentItemId) return { id: args.parentItemId };
  let parentPath = "";
  if (args.destinationParentPreset) {
    parentPath = resolvePresetPath(args, {
      pathField: "destinationParentPath",
      presetField: "destinationParentPreset",
      relativeField: "destinationParentRelativePath"
    });
  } else if (args.parentPreset) {
    parentPath = resolvePresetPath(args, {
      pathField: "parentPath",
      presetField: "parentPreset",
      relativeField: "parentRelativePath"
    });
  } else {
    parentPath = args.destinationParentPath ?? args.parentPath ?? "";
  }
  const raw = cleanPath(parentPath) === "" ? await graph("/me/drive/root") : await getRawInfo({ path: parentPath });
  if (!raw.folder && !raw.root) throw new Error(`Destination parent is not a folder: ${raw.name}`);
  return { id: raw.id, path: raw.parentReference?.path, name: raw.name };
}

async function readText(args = {}) {
  const info = await getInfo(args);
  const maxBytes = clampInteger(args.maxBytes, textFileLimit, 1, maxTextFileReadLimit);
  assertTextReadable(info, args);
  if (info.size && info.size > maxBytes) {
    throw new Error(`File is ${info.size} bytes, above maxBytes ${maxBytes}. Use onedrive_download instead.`);
  }
  const limited = await graphLimitedBuffer(contentPath(args), maxBytes);
  if (limited.truncated) {
    throw new Error(`Downloaded content is above maxBytes ${maxBytes}. Use onedrive_download instead.`);
  }
  const buffer = limited.buffer;
  if (args.force !== true) assertNoBinaryNulls(buffer);
  return { item: info, content: buffer.toString("utf8") };
}

async function download(args = {}) {
  if (args.localPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.localPath), "Download", args);
  const info = await getInfo(args);
  return await downloadResolvedItem(info, args);
}

async function downloadResolvedItem(info, args = {}) {
  const preferredTarget = args.localPath ? resolve(args.localPath) : join(downloadRoot, info.name || basename(cleanPath(args.path || args.itemId || "download")));
  await assertNotLocalOneDriveSyncPathForWrite(preferredTarget, "Download", args);
  const reservation = await reserveLocalDestination(preferredTarget, {
    overwrite: args.overwrite === true,
    allowAlternate: !args.localPath || args._allowAlternateLocalPath === true
  });
  const target = reservation.path;
  const contentArgs = info.id ? { itemId: info.id } : args;
  try {
    const downloaded = await graphDownloadToFile(contentPath(contentArgs), target, { reserved: reservation.reserved });
    return { item: info, localPath: target, bytesWritten: downloaded.bytesWritten };
  } catch (error) {
    if (reservation.reserved) await rm(target, { force: true });
    throw error;
  }
}

function officePackageKindFromName(name = "") {
  const extension = extname(name).toLowerCase();
  if ([".docx", ".docm"].includes(extension)) return "word";
  if ([".xlsx", ".xlsm"].includes(extension)) return "excel";
  if ([".pptx", ".pptm", ".ppsx"].includes(extension)) return "powerpoint";
  return null;
}

function officeRuntimeStatus() {
  let pythonAvailable = false;
  let pythonVersion = null;
  let helperAvailable = false;
  let error = null;
  try {
    pythonVersion = execFileSync(officePythonPath, ["--version"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
      timeout: 5000
    }).trim();
    pythonAvailable = true;
  } catch (runtimeError) {
    error = safeToolErrorMessage(runtimeError);
  }
  try {
    helperAvailable = readFileSync(officeHelperPath).length > 0;
  } catch (helperError) {
    error ||= safeToolErrorMessage(helperError);
  }
  return { pythonAvailable, pythonVersion, pythonPath: officePythonPath, helperAvailable, helperPath: officeHelperPath, error };
}

async function officeCapabilities() {
  const runtime = officeRuntimeStatus();
  let drive = null;
  let driveError = null;
  try {
    drive = await graph("/me/drive");
  } catch (error) {
    driveError = safeToolErrorMessage(error);
  }
  const driveType = drive?.driveType || null;
  return {
    runtime,
    account: {
      driveType,
      driveName: drive?.name || null,
      checked: Boolean(drive),
      error: driveError
    },
    backends: {
      openXml: {
        available: runtime.pythonAvailable && runtime.helperAvailable,
        formats: [".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".ppsx"],
        readOnlyToolsReady: true,
        mutationToolsReady: true,
        operations: {
          word: ["replaceText", "setParagraphText", "setParagraphStyle", "insertParagraph", "insertTable", "setTableCell", "setContentControlText", "addHyperlink", "addComment"],
          excel: ["setCell", "setFormula", "setRange", "clearRange", "setStyle", "setNumberFormat", "addConditionalFormat", "setDataValidation", "freezePanes", "setColumnWidth", "renameSheet", "setDefinedName", "recalculate", "addTableRow", "setTableTotals", "createChart", "updateChart"],
          powerpoint: ["replaceText", "setShapeText", "setShapeGeometry", "setTextStyle", "addTextBox", "deleteShape", "replaceImage", "setTableCell", "setNotes", "duplicateSlide", "deleteSlide", "moveSlide"]
        },
        notes: [
          "Encrypted and legacy binary Office files are refused.",
          "Macro-enabled packages are inspectable but live mutation will require explicit macro-preservation safeguards.",
          "Digitally signed packages are detected and mutation is always refused."
        ]
      },
      graphExcel: {
        availableForAccount: Boolean(drive && driveType !== "personal"),
        driveType,
        formats: [".xlsx"],
        operations: ["setCell", "setFormula", "setRange", "clearRange", "renameSheet", "addTableRow", "createChart", "updateChart"],
        limitation: driveType === "personal"
          ? "Microsoft Graph workbook APIs do not support OneDrive Consumer workbooks; use the Open XML backend."
          : "Graph workbook sessions will be used for supported business, SharePoint, or group-drive .xlsx files."
      },
      officeJs: {
        available: false,
        requiresCompanionAddIn: true,
        protocolVersion: "codex-office-companion/1",
        companionVersion: "1.1.0",
        requirementSets: { word: "WordApi 1.3", excel: "ExcelApi 1.7", powerpoint: "PowerPointApi 1.5" },
        operations: {
          word: ["replaceText", "setSelectedText", "insertParagraph"],
          excel: ["setRange", "clearRange", "formatRange"],
          powerpoint: ["setSelectedText", "setSelectedTextStyle", "deleteSelectedShapes"]
        },
        transport: { manualPaste: true, remoteCommands: false, telemetry: false },
        note: "Office.js commands require an active Word, Excel, or PowerPoint client and the separately deployed companion add-in; available remains false because the MCP server cannot attest to an active task pane."
      }
    }
  };
}

async function runOfficeHelper(request, options = {}) {
  const runtime = officeRuntimeStatus();
  if (!runtime.pythonAvailable || !runtime.helperAvailable) {
    throw new Error(`Office Open XML runtime is unavailable: ${runtime.error || "Python/helper not found"}`);
  }
  const timeoutMs = clampInteger(options.timeoutMs, 60_000, 1000, 5 * 60_000);
  const maxOutputBytes = clampInteger(options.maxOutputBytes, 20 * 1024 * 1024, 1024, 50 * 1024 * 1024);
  await ensurePrivateDirectory(storageRoot);
  const pythonCacheRoot = join(storageRoot, "pycache");
  await ensurePrivateDirectory(pythonCacheRoot);
  return await new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(officePythonPath, [officeHelperPath], {
      cwd: pluginRoot,
      stdio: ["pipe", "pipe", "pipe"],
      env: { ...process.env, PYTHONPYCACHEPREFIX: pythonCacheRoot }
    });
    let stdout = Buffer.alloc(0);
    let stderr = Buffer.alloc(0);
    let settled = false;
    let timer = null;
    const finish = (error, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (error) rejectPromise(error);
      else resolvePromise(value);
    };
    const appendBounded = (current, chunk, label) => {
      const next = Buffer.concat([current, chunk]);
      if (next.length > maxOutputBytes) {
        child.kill("SIGKILL");
        finish(new Error(`Office helper ${label} exceeded ${maxOutputBytes} bytes.`));
      }
      return next;
    };
    child.stdout.on("data", (chunk) => { stdout = appendBounded(stdout, chunk, "stdout"); });
    child.stderr.on("data", (chunk) => { stderr = appendBounded(stderr, chunk, "stderr"); });
    child.on("error", (error) => finish(new Error(`Could not start Office helper: ${error.message}`)));
    child.on("close", (code) => {
      let parsed = null;
      try {
        parsed = JSON.parse(stdout.toString("utf8") || "{}");
      } catch (error) {
        finish(new Error(`Office helper returned invalid JSON: ${error.message}. ${stderr.toString("utf8").trim()}`));
        return;
      }
      if (code !== 0 || parsed.ok !== true) {
        finish(new Error(parsed.error || stderr.toString("utf8").trim() || `Office helper exited with code ${code}.`));
        return;
      }
      finish(null, parsed.value);
    });
    timer = setTimeout(() => {
      child.kill("SIGKILL");
      finish(new Error(`Office helper timed out after ${timeoutMs}ms.`));
    }, timeoutMs);
    child.stdin.on("error", (error) => finish(new Error(`Could not send request to Office helper: ${error.message}`)));
    child.stdin.end(JSON.stringify(request));
  });
}

async function inspectRemoteOfficePackage(args = {}, expectedKind = null, action = "inspect") {
  const info = await getInfo(args);
  if (info.folder) throw new Error(`Item is a folder, not an Office document: ${info.name}`);
  const detectedKind = officePackageKindFromName(info.name);
  if (!detectedKind) {
    throw new Error(`Office content tools require an Open XML package (.docx, .xlsx, or .pptx family). Got ${info.name || "unnamed item"}.`);
  }
  if (expectedKind && detectedKind !== expectedKind) {
    throw new Error(`Expected a ${expectedKind} Open XML package. Got ${info.name}.`);
  }
  if (Number(info.size || 0) > maxOfficePackageBytes) {
    throw new Error(`Office package is ${info.size} bytes, above the ${maxOfficePackageBytes}-byte inspection limit.`);
  }
  const transactionRoot = join(officeEditingRoot, `inspect-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const localPath = join(transactionRoot, basename(info.name));
  try {
    await downloadResolvedItem(info, { localPath, overwrite: false });
    const value = await runOfficeHelper({
      ...args,
      action,
      inputPath: localPath,
      kind: expectedKind || args.expectedKind || detectedKind
    });
    return {
      item: info,
      backend: "openxml",
      ...value,
      package: value.package ? { ...value.package, path: undefined } : undefined
    };
  } finally {
    await rm(transactionRoot, { recursive: true, force: true });
  }
}

function officeEditPreviewProof(rawItem, kind, operations, editResult, args = {}, backend = "openxml") {
  return {
    item: { id: rawItem.id, name: rawItem.name, eTag: rawItem.eTag, cTag: rawItem.cTag },
    kind,
    backend,
    operations,
    changeCount: editResult.changeCount,
    outputFingerprint: editResult.validation?.package?.fingerprint,
    allowMacros: args.allowMacros === true,
    allowSignedPackage: args.allowSignedPackage === true
  };
}

const excelGraphOperationTypes = new Set(["setCell", "setFormula", "setRange", "clearRange", "renameSheet", "addTableRow", "createChart", "updateChart"]);
const excelGraphOnlyOperationTypes = new Set(["createChart", "updateChart"]);

async function resolveExcelOfficeBackend(rawItem, args = {}) {
  if (args.backend === "openxml") return { backend: "openxml", driveType: null, reason: "explicit" };
  if (extname(rawItem.name || "").toLowerCase() !== ".xlsx") {
    if (args.backend === "graph") throw new Error("The Graph Excel backend supports .xlsx only; use openxml for this workbook.");
    return { backend: "openxml", driveType: null, reason: "format" };
  }
  const unsupported = (args.operations || []).map((entry) => entry?.type).filter((type) => !excelGraphOperationTypes.has(type));
  if (unsupported.length) {
    if (args.backend === "graph") throw new Error(`The Graph Excel backend does not support these operations: ${[...new Set(unsupported)].join(", ")}.`);
    return { backend: "openxml", driveType: null, reason: "operation" };
  }
  const target = excelWorkbookTarget(rawItem);
  const driveId = target.driveId;
  let drive;
  try {
    drive = driveId ? await graph(`/drives/${encodeURIComponent(driveId)}?$select=id,driveType,name`) : await graph("/me/drive?$select=id,driveType,name");
  } catch (error) {
    if (args.backend === "graph") throw new Error(`Could not verify the workbook drive for Graph Excel: ${safeToolErrorMessage(error)}`);
    return { backend: "openxml", driveType: null, reason: "drive-check-failed" };
  }
  if (drive.driveType === "personal") {
    if (args.backend === "graph") throw new Error("Microsoft Graph workbook APIs do not support OneDrive Consumer workbooks; use backend: openxml.");
    return { backend: "openxml", driveType: drive.driveType, driveId: drive.id, reason: "consumer" };
  }
  return { backend: "graph", driveType: drive.driveType, driveId: drive.id || driveId, reason: args.backend === "graph" ? "explicit" : "supported-drive" };
}

function excelWorkbookTarget(rawItem) {
  const remoteItem = rawItem.remoteItem;
  if (remoteItem?.id && remoteItem.parentReference?.driveId) {
    return { driveId: remoteItem.parentReference.driveId, itemId: remoteItem.id };
  }
  return {
    driveId: rawItem.parentReference?.driveId || remoteItem?.parentReference?.driveId || null,
    itemId: rawItem.id
  };
}

function excelWorkbookBase(rawItem, resolvedBackend) {
  const target = excelWorkbookTarget(rawItem);
  const driveId = resolvedBackend.driveId || target.driveId;
  return driveId
    ? `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(target.itemId)}/workbook`
    : `/me/drive/items/${encodeURIComponent(target.itemId)}/workbook`;
}

function excelGraphRangePath(base, operation) {
  const sheet = String(operation.sheet || "");
  const address = String(operation.address || "").toUpperCase();
  if (!sheet || !/^[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?$/.test(address)) {
    throw new Error("Graph Excel range operations require a worksheet name and a bounded A1 address.");
  }
  const escapedAddress = address.replaceAll("'", "''");
  return `${base}/worksheets/${encodeURIComponent(sheet)}/range(address='${escapedAddress}')`;
}

async function applyExcelGraphOperations(rawItem, resolvedBackend, operations) {
  const base = excelWorkbookBase(rawItem, resolvedBackend);
  const created = await createExcelWorkbookSession(base);
  const sessionId = created?.id;
  if (!sessionId) throw new Error("Microsoft Graph did not return an Excel workbook session ID.");
  const headers = { "workbook-session-id": sessionId };
  let writeStarted = false;
  let sessionClosed = false;
  try {
    for (const operation of operations) {
      if (operation.type === "addTableRow") {
        writeStarted = true;
        await graph(`${base}/tables/${encodeURIComponent(String(operation.table || ""))}/rows/add`, {
          method: "POST", headers, body: JSON.stringify({ values: operation.values, index: operation.index ?? null }), maxRetries: 0
        });
        continue;
      }
      if (operation.type === "createChart") {
        writeStarted = true;
        await graph(`${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}/charts/add`, {
          method: "POST", headers, body: JSON.stringify({ type: operation.chartType, sourceData: operation.sourceData, seriesBy: operation.seriesBy || "Auto" }), maxRetries: 0
        });
        continue;
      }
      if (operation.type === "updateChart") {
        const chartBase = `${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}/charts/${encodeURIComponent(String(operation.chart || ""))}`;
        const chartPatch = Object.fromEntries(["name", "height", "width", "left", "top"].filter((key) => operation[key] !== undefined).map((key) => [key, operation[key]]));
        if (Object.keys(chartPatch).length) {
          writeStarted = true;
          await graph(chartBase, { method: "PATCH", headers, body: JSON.stringify(chartPatch), maxRetries: 0 });
        }
        if (operation.titleText !== undefined) {
          writeStarted = true;
          await graph(`${chartBase}/title`, { method: "PATCH", headers, body: JSON.stringify({ text: operation.titleText }), maxRetries: 0 });
        }
        if (operation.sourceData !== undefined) {
          writeStarted = true;
          await graph(`${chartBase}/setData`, { method: "POST", headers, body: JSON.stringify({ sourceData: operation.sourceData, seriesBy: operation.seriesBy || "Auto" }), maxRetries: 0 });
        }
        continue;
      }
      if (operation.type === "renameSheet") {
        writeStarted = true;
        await graph(`${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}`, {
          method: "PATCH", headers, body: JSON.stringify({ name: operation.newName }), maxRetries: 0
        });
        continue;
      }
      const rangePath = excelGraphRangePath(base, operation);
      if (operation.type === "clearRange") {
        writeStarted = true;
        await graph(`${rangePath}/clear`, {
          method: "POST", headers, body: JSON.stringify({ applyTo: operation.format ? "All" : "Contents" }), maxRetries: 0
        });
        continue;
      }
      const body = operation.type === "setCell"
        ? { values: [[operation.value]] }
        : operation.type === "setFormula"
          ? { formulas: [[operation.formula]], ...(operation.value !== undefined ? { values: [[operation.value]] } : {}) }
          : { ...(operation.values !== undefined ? { values: operation.values } : {}), ...(operation.formulas !== undefined ? { formulas: operation.formulas } : {}) };
      writeStarted = true;
      await graph(rangePath, { method: "PATCH", headers, body: JSON.stringify(body), maxRetries: 0 });
    }
  } catch (error) {
    const suffix = writeStarted ? " A persistent workbook session had begun; the request was not replayed." : "";
    throw new Error(`Graph Excel update failed: ${safeToolErrorMessage(error)}${suffix}`);
  } finally {
    try {
      await graph(`${base}/closeSession`, { method: "POST", headers, maxRetries: 0 });
      sessionClosed = true;
    } catch (error) {
      toolCallContext.getStore()?.localWarnings?.push({ operation: "Graph Excel session close", error: safeToolErrorMessage(error) });
    }
  }
  return { sessionClosed, operationCount: operations.length };
}

function officeSemanticDiff(kind, changes = []) {
  const operationCounts = {};
  const affectedParts = new Set();
  const affectedObjects = [];
  for (const change of changes) {
    const operation = change.operation || "unknown";
    operationCounts[operation] = (operationCounts[operation] || 0) + 1;
    if (change.part) affectedParts.add(change.part);
    const selector = Object.fromEntries(["part", "paragraphIndex", "tableIndex", "rowIndex", "columnIndex", "contentControlIndex", "sheet", "address", "name", "slideIndex", "shapeId", "commentId"]
      .filter((key) => change[key] !== undefined).map((key) => [key, change[key]]));
    if (Object.keys(selector).length) affectedObjects.push({ operation, ...selector });
  }
  return { kind, changeCount: changes.length, operationCounts, affectedParts: [...affectedParts].sort(), affectedObjects: affectedObjects.slice(0, 200), affectedObjectsTruncated: affectedObjects.length > 200 };
}

function officeInspectionSummary(kind, value = {}) {
  if (kind === "word") return { paragraphCount: value.paragraphCount || 0, tableCount: value.tableCount || 0, contentControlCount: value.contentControlCount || 0, commentCount: value.commentCount || 0 };
  if (kind === "excel") return { sheetCount: value.sheetCount || 0, sheets: (value.sheets || []).map((sheet) => ({ name: sheet.name, cellCount: sheet.cellCount ?? sheet.cells?.length ?? 0 })) };
  return { slideCount: value.slideCount || 0, slides: (value.slides || []).map((slide) => ({ index: slide.index, shapeCount: slide.shapeCount ?? slide.shapes?.length ?? 0 })) };
}

function compareOfficeInspections(kind, before = {}, after = {}, maxChanges = 100) {
  const changes = [];
  let totalChanges = 0;
  const add = (change) => { totalChanges += 1; if (changes.length < maxChanges) changes.push(change); };
  if (kind === "word") {
    const paragraphs = (value) => new Map((value.paragraphs || []).map((entry) => [`${entry.part}:${entry.index}`, entry]));
    const leftMap = paragraphs(before), rightMap = paragraphs(after);
    for (const key of new Set([...leftMap.keys(), ...rightMap.keys()])) {
      const left = leftMap.get(key), right = rightMap.get(key);
      if (left?.text !== right?.text || left?.style !== right?.style) add({ objectType: "paragraph", key, before: left ? { text: left.text, style: left.style } : null, after: right ? { text: right.text, style: right.style } : null });
    }
  } else if (kind === "excel") {
    const cells = (value) => new Map((value.sheets || []).flatMap((sheet) => (sheet.cells || []).map((cell) => [`${sheet.name}!${cell.address}`, cell])));
    const leftMap = cells(before), rightMap = cells(after);
    for (const key of new Set([...leftMap.keys(), ...rightMap.keys()])) {
      const left = leftMap.get(key), right = rightMap.get(key);
      if (JSON.stringify([left?.value, left?.formula, left?.styleIndex]) !== JSON.stringify([right?.value, right?.formula, right?.styleIndex])) add({ objectType: "cell", key, before: left ? { value: left.value, formula: left.formula, styleIndex: left.styleIndex } : null, after: right ? { value: right.value, formula: right.formula, styleIndex: right.styleIndex } : null });
    }
  } else {
    const shapes = (value) => new Map((value.slides || []).flatMap((slide) => (slide.shapes || []).map((shape) => [`${slide.index}:${shape.id}`, shape])));
    const leftMap = shapes(before), rightMap = shapes(after);
    for (const key of new Set([...leftMap.keys(), ...rightMap.keys()])) {
      const left = leftMap.get(key), right = rightMap.get(key);
      if (JSON.stringify([left?.text, left?.geometry]) !== JSON.stringify([right?.text, right?.geometry])) add({ objectType: "shape", key, before: left ? { text: left.text, geometry: left.geometry } : null, after: right ? { text: right.text, geometry: right.geometry } : null });
    }
  }
  const beforeSummary = officeInspectionSummary(kind, before), afterSummary = officeInspectionSummary(kind, after);
  return { kind, sameSemantics: totalChanges === 0 && JSON.stringify(beforeSummary) === JSON.stringify(afterSummary), before: beforeSummary, after: afterSummary, changeCount: totalChanges, changes, truncated: totalChanges > changes.length };
}

function assertOfficeBackupId(backupId) {
  const value = String(backupId || "");
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)) throw new Error("backupId is invalid.");
  return value.toLowerCase();
}

function officeBackupManifestPath(backupId) {
  return join(backupRoot, `office-${assertOfficeBackupId(backupId)}.json`);
}

async function persistOfficeBackup(localPath, rawItem, kind, fingerprint, reason = "edit") {
  await ensurePrivateDirectory(backupRoot);
  const backupId = randomUUID();
  const fileName = `office-${backupId}-${basename(rawItem.name || `${kind}.bin`)}`;
  const backupPath = join(backupRoot, fileName);
  await copyFile(localPath, backupPath);
  await chmod(backupPath, 0o600);
  const manifest = { version: 1, backupId, createdAt: new Date().toISOString(), reason, kind, fileName, bytes: (await stat(backupPath)).size, fingerprint: fingerprint || null, item: { id: rawItem.id, name: rawItem.name, remotePath: itemRemotePath(rawItem) || null, eTag: rawItem.eTag || null, cTag: rawItem.cTag || null, size: rawItem.size ?? null, lastModifiedDateTime: rawItem.lastModifiedDateTime || null } };
  await writePrivateFileAtomic(officeBackupManifestPath(backupId), JSON.stringify(manifest));
  return { backupId, createdAt: manifest.createdAt, kind, item: manifest.item, bytes: manifest.bytes, localPath: backupPath, reason };
}

async function loadOfficeBackup(backupId) {
  const id = assertOfficeBackupId(backupId);
  const manifestPath = officeBackupManifestPath(id);
  let manifest;
  try {
    const manifestInfo = await lstat(manifestPath);
    if (!manifestInfo.isFile() || manifestInfo.isSymbolicLink()) throw new Error("unsafe manifest");
    manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  }
  catch { throw new Error(`Office backup ${id} was not found or its manifest is invalid.`); }
  if (manifest.version !== 1 || manifest.backupId !== id || !manifest.item?.id || !["word", "excel", "powerpoint"].includes(manifest.kind)) throw new Error(`Office backup ${id} has an invalid manifest.`);
  if (!manifest.fileName || basename(manifest.fileName) !== manifest.fileName || !manifest.fileName.startsWith(`office-${id}-`)) throw new Error(`Office backup ${id} has an unsafe file reference.`);
  const localPath = join(backupRoot, manifest.fileName);
  const info = await lstat(localPath);
  if (!info.isFile() || info.isSymbolicLink() || info.size > maxOfficePackageBytes) throw new Error(`Office backup ${id} is missing, unsafe, or exceeds the Office size limit.`);
  const [resolvedRoot, resolvedFile] = await Promise.all([realpath(backupRoot), realpath(localPath)]);
  if (dirname(resolvedFile) !== resolvedRoot) throw new Error(`Office backup ${id} resolves outside managed backup storage.`);
  return { manifest, localPath };
}

async function officeBackups(args = {}) {
  await ensurePrivateDirectory(backupRoot);
  const entries = await readdir(backupRoot, { withFileTypes: true });
  const items = [];
  for (const entry of entries) {
    const match = entry.isFile() && entry.name.match(/^office-([0-9a-f-]{36})\.json$/i);
    if (!match) continue;
    try {
      const { manifest } = await loadOfficeBackup(match[1]);
      if (args.itemId && manifest.item.id !== args.itemId) continue;
      if (args.kind && manifest.kind !== args.kind) continue;
      items.push({ backupId: manifest.backupId, createdAt: manifest.createdAt, reason: manifest.reason, kind: manifest.kind, bytes: manifest.bytes, fingerprint: manifest.fingerprint, item: manifest.item });
    } catch { /* Ignore malformed legacy/unmanaged entries. */ }
  }
  items.sort((left, right) => String(right.createdAt).localeCompare(String(left.createdAt)));
  const limit = clampInteger(args.limit, 100, 1, 500);
  return { count: Math.min(items.length, limit), total: items.length, items: items.slice(0, limit), truncated: items.length > limit };
}

async function inspectLocalOffice(localPath, kind) {
  return await runOfficeHelper({ action: "inspect", inputPath: localPath, kind, maxParagraphs: 10000, maxCells: 50000, maxSlides: 5000 });
}

async function officeCompareBackup(args = {}) {
  const { manifest, localPath } = await loadOfficeBackup(args.backupId);
  const current = await getRawInfo({ itemId: manifest.item.id });
  if (officePackageKindFromName(current.name) !== manifest.kind) throw new Error("Current remote item no longer has the backup's Office document kind.");
  const transactionRoot = join(officeEditingRoot, `compare-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const currentPath = join(transactionRoot, basename(current.name));
  try {
    await downloadResolvedItem(simplifyItem(current), { localPath: currentPath, overwrite: false });
    const [backupInspection, currentInspection] = await Promise.all([inspectLocalOffice(localPath, manifest.kind), inspectLocalOffice(currentPath, manifest.kind)]);
    return { backup: { backupId: manifest.backupId, createdAt: manifest.createdAt, kind: manifest.kind, bytes: manifest.bytes, fingerprint: backupInspection.package?.fingerprint, item: manifest.item }, current: { item: simplifyItem(current), fingerprint: currentInspection.package?.fingerprint }, sameContent: backupInspection.package?.fingerprint === currentInspection.package?.fingerprint, semanticDiff: compareOfficeInspections(manifest.kind, backupInspection, currentInspection, clampInteger(args.maxChanges, 100, 1, 500)) };
  } finally { await rm(transactionRoot, { recursive: true, force: true }); }
}

async function officeRestoreBackup(args = {}) {
  const { manifest, localPath } = await loadOfficeBackup(args.backupId);
  const current = await getRawInfo({ itemId: manifest.item.id });
  const backupValidation = await runOfficeHelper({ action: "validate", inputPath: localPath, kind: manifest.kind });
  const comparison = await officeCompareBackup({ backupId: manifest.backupId, maxChanges: 100 });
  const preview = { dryRun: args.dryRun !== false, confirmed: args.confirmed === true, wouldRestore: { backupId: manifest.backupId, createdAt: manifest.createdAt, kind: manifest.kind, originalItem: manifest.item, currentItem: simplifyItem(current), backupFingerprint: backupValidation.package?.fingerprint }, sameContent: comparison.sameContent, semanticDiff: comparison.semanticDiff };
  const proof = { backupId: manifest.backupId, itemId: current.id, currentETag: current.eTag, backupFingerprint: backupValidation.package?.fingerprint };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_office_restore_backup", proof);
  if (args.confirmed !== true) return { ...preview, requiredToRestore: "Set dryRun: false and confirmed: true after reviewing the Office backup preview." };
  if (!args.expectedId || args.expectedId !== current.id || args.expectedId !== manifest.item.id) return { ...preview, dryRun: false, requiredToRestore: "Provide expectedId matching the backup's original stable item ID." };
  if (!args.expectedETag || args.expectedETag !== current.eTag) return { ...preview, dryRun: false, requiredToRestore: "Provide expectedETag matching the current remote item eTag." };
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_office_restore_backup", proof, args.previewToken, "requiredToRestore");
  if (tokenRequired) return tokenRequired;
  const transactionRoot = join(officeEditingRoot, `restore-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  let rollbackBackup = null, verificationIncomplete = false, afterRaw = current;
  try {
    const currentPath = join(transactionRoot, `current-${basename(current.name)}`);
    await downloadResolvedItem(simplifyItem(current), { localPath: currentPath, overwrite: false });
    const currentValidation = await runOfficeHelper({ action: "validate", inputPath: currentPath, kind: manifest.kind });
    rollbackBackup = await persistOfficeBackup(currentPath, current, manifest.kind, currentValidation.package?.fingerprint, "pre-restore");
    const uploaded = await upload({ localPath, itemId: current.id, remotePath: itemRemotePath(current), conflictBehavior: "replace", ifMatch: current.eTag, auditTool: "onedrive_office_restore_backup", skipAudit: true });
    afterRaw = uploaded.item ? { ...current, ...uploaded.item } : current;
    try { afterRaw = await getRawInfo({ itemId: current.id }); } catch (error) { verificationIncomplete = true; recordLocalWarning("Office backup restore metadata verification", error); }
    let remoteValidation = null;
    if (args.verify !== false) {
      try {
        const verifyPath = join(transactionRoot, `verify-${basename(current.name)}`);
        await downloadResolvedItem(simplifyItem(afterRaw), { localPath: verifyPath, overwrite: false });
        remoteValidation = await runOfficeHelper({ action: "validate", inputPath: verifyPath, kind: manifest.kind });
        if (remoteValidation.package?.fingerprint !== backupValidation.package?.fingerprint) { verificationIncomplete = true; recordLocalWarning("Office backup restore fingerprint verification", new Error("Restored package fingerprint differs from the selected backup.")); }
      } catch (error) { verificationIncomplete = true; recordLocalWarning("Office backup restore validation", error); }
    }
    await writeMutationAudit("onedrive_office_restore_backup", { status: "success", target: itemAuditSummary(current), before: itemAuditSummary(current), after: itemAuditSummary(afterRaw), backupId: manifest.backupId, rollbackBackupId: rollbackBackup.backupId, verificationIncomplete });
    return { dryRun: false, confirmed: true, restoredBackupId: manifest.backupId, item: simplifyItem(afterRaw), rollbackBackup, backupValidation, remoteValidation, verificationIncomplete, uploaded };
  } catch (error) {
    await writeMutationAudit("onedrive_office_restore_backup", { status: "failed", target: itemAuditSummary(current), backupId: manifest.backupId, rollbackBackupId: rollbackBackup?.backupId, error: safeErrorInfo(error) });
    throw error;
  } finally { await rm(transactionRoot, { recursive: true, force: true }); }
}

function trustedExcelSessionOperationUrl(value, label) {
  if (!value) throw new Error(`Microsoft Graph returned an asynchronous Excel session response without ${label}.`);
  try {
    return graphUrl(value);
  } catch (error) {
    throw new Error(`Microsoft Graph returned an untrusted Excel session ${label}: ${safeToolErrorMessage(error)}`);
  }
}

async function createExcelWorkbookSession(base) {
  // Microsoft documents 504 as safe to retry specifically for createSession.
  const requestOptions = {
    method: "POST",
    body: JSON.stringify({ persistChanges: true }),
    headers: { Prefer: "respond-async" },
    maxRetries: 0,
    returnResponse: true
  };
  let response = await graph(`${base}/createSession`, requestOptions);
  if (response.status === 504) response = await graph(`${base}/createSession`, requestOptions);
  if (!response.ok) {
    throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Excel session creation failed" });
  }
  if (response.status !== 202) return response.body;

  const operationUrl = trustedExcelSessionOperationUrl(response.headers.get("location"), "Location URL");
  const timeoutMs = clampInteger(process.env.ONEDRIVE_EXCEL_SESSION_TIMEOUT_MS, 60_000, 1000, 5 * 60_000);
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const operation = await graph(operationUrl, { returnResponse: true, maxRetries: 3 });
    if (!operation.ok) {
      throw microsoftGraphError(operation.body, { headers: operation.headers, status: operation.status, statusText: "Excel session operation failed" });
    }
    const status = String(operation.body?.status || "");
    if (status === "succeeded") {
      const resourceLocation = operation.body?.resourceLocation || operation.headers.get("resource-location");
      const resourceUrl = trustedExcelSessionOperationUrl(resourceLocation, "resourceLocation URL");
      return await graph(resourceUrl, { maxRetries: 3 });
    }
    if (status === "failed") {
      const detail = operation.body?.error?.message || operation.body?.error?.code || "the operation reported failure";
      throw new Error(`Microsoft Graph Excel session creation failed: ${detail}`);
    }
    if (status && !["notStarted", "running", "inProgress"].includes(status)) {
      throw new Error(`Microsoft Graph returned an unknown Excel session operation status: ${status}`);
    }
    const configuredDelay = clampInteger(process.env.ONEDRIVE_EXCEL_SESSION_POLL_MS, 2000, 0, 10_000);
    await sleep(operation.headers.get("retry-after") ? retryDelayMs(operation, 0) : configuredDelay);
  }
  throw new Error(`Microsoft Graph Excel session creation did not complete within ${timeoutMs}ms.`);
}

async function officeBatchUpdate(args = {}, kind, toolName) {
  const rawItem = await getRawInfo(args);
  if (rawItem.folder) throw new Error(`Item is a folder, not an Office document: ${rawItem.name}`);
  assertExpectedItem(rawItem, args, `${kind} batch update`);
  const detectedKind = officePackageKindFromName(rawItem.name);
  if (detectedKind !== kind) throw new Error(`Expected a ${kind} Open XML package. Got ${rawItem.name}.`);
  if (Number(rawItem.size || 0) > maxOfficePackageBytes) {
    throw new Error(`Office package is ${rawItem.size} bytes, above the ${maxOfficePackageBytes}-byte editing limit.`);
  }
  const backendResolution = kind === "excel" ? await resolveExcelOfficeBackend(rawItem, args) : { backend: "openxml" };
  const selectedBackend = backendResolution.backend;
  if (args.dryRun === false && args.confirmed !== true) {
    return { dryRun: false, confirmed: false, requiredToUpdate: "Set dryRun: false and confirmed: true after reviewing the Office edit preview." };
  }
  if (args.dryRun === false && !hasExpectedIdentity(args)) {
    return { dryRun: false, confirmed: true, requiredToUpdate: "Provide expectedName or expectedId for live Office edits." };
  }

  const transactionRoot = join(officeEditingRoot, `edit-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const sourcePath = join(transactionRoot, `source-${basename(rawItem.name)}`);
  const editedPath = join(transactionRoot, `edited-${basename(rawItem.name)}`);
  let editResult;
  let backup = null;
  try {
    await downloadResolvedItem(simplifyItem(rawItem), { localPath: sourcePath, overwrite: false });
    const helperOperations = kind === "excel" && selectedBackend === "graph"
      ? args.operations.filter((operation) => !excelGraphOnlyOperationTypes.has(operation.type))
      : args.operations;
    const graphOnlyChanges = kind === "excel" && selectedBackend === "graph"
      ? args.operations.filter((operation) => excelGraphOnlyOperationTypes.has(operation.type)).map((operation) => ({
          operation: operation.type,
          sheet: operation.sheet,
          table: operation.table,
          chart: operation.chart,
          after: Object.fromEntries(Object.entries(operation).filter(([key]) => key !== "type"))
        }))
      : [];
    if (helperOperations.length) {
      editResult = await runOfficeHelper({
        action: "edit",
        inputPath: sourcePath,
        outputPath: editedPath,
        kind,
        operations: helperOperations,
        ...(kind === "word" ? { trackedChanges: args.trackedChanges || "refuse" } : {}),
        allowMacros: args.allowMacros === true,
        allowSignedPackage: args.allowSignedPackage === true
      }, { timeoutMs: 120_000 });
    } else {
      editResult = { kind, changes: [], changeCount: 0, validation: await runOfficeHelper({ action: "validate", inputPath: sourcePath, kind }) };
    }
    if (graphOnlyChanges.length) {
      editResult.changes = [...(editResult.changes || []), ...graphOnlyChanges];
      editResult.changeCount = (editResult.changeCount || 0) + graphOnlyChanges.length;
    }
    if (!editResult.changeCount) {
      return {
        dryRun: args.dryRun !== false,
        confirmed: args.confirmed === true,
        item: simplifyItem(rawItem),
        backend: "openxml",
        changes: [],
        changeCount: 0,
        semanticDiff: officeSemanticDiff(kind, []),
        noChanges: true,
        requiredToUpdate: "No requested edit matched the document. Re-read the document and correct the operation anchors."
      };
    }
    const proof = officeEditPreviewProof(rawItem, kind, args.operations, editResult, args, selectedBackend);
    const preview = {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      item: simplifyItem(rawItem),
      backend: selectedBackend,
      previewBackend: selectedBackend === "graph" ? "openxml" : undefined,
      changes: editResult.changes,
      changeCount: editResult.changeCount,
      semanticDiff: officeSemanticDiff(kind, editResult.changes),
      validation: editResult.validation
    };
    if (args.dryRun !== false) return previewWithToken(preview, toolName, proof);
    const tokenRequired = previewTokenRequiredResult(preview, toolName, proof, args.previewToken, "requiredToUpdate");
    if (tokenRequired) return tokenRequired;

    if (args.createBackup !== false) {
      const sourceValidation = await runOfficeHelper({ action: "validate", inputPath: sourcePath, kind });
      backup = await persistOfficeBackup(sourcePath, rawItem, kind, sourceValidation.package?.fingerprint, "edit");
    }

    const uploaded = selectedBackend === "graph"
      ? { item: simplifyItem(rawItem), ...(await applyExcelGraphOperations(rawItem, backendResolution, args.operations)), uploadMode: "graph-workbook-session" }
      : await upload({
        localPath: editedPath,
        itemId: rawItem.id,
        remotePath: itemRemotePath(rawItem),
        conflictBehavior: "replace",
        ifMatch: rawItem.eTag,
        auditTool: toolName,
        skipAudit: true
      });
    let afterRaw = uploaded.item ? { ...rawItem, ...uploaded.item } : rawItem;
    let remoteValidation = null;
    let verificationIncomplete = false;
    try {
      afterRaw = await getRawInfo({ itemId: rawItem.id });
    } catch (error) {
      verificationIncomplete = true;
      toolCallContext.getStore()?.localWarnings?.push({ operation: "Office post-commit metadata verification", error: safeToolErrorMessage(error) });
    }
    if (args.verify !== false) {
      try {
        const verifyPath = join(transactionRoot, `verify-${basename(rawItem.name)}`);
        await downloadResolvedItem(simplifyItem(afterRaw), { localPath: verifyPath, overwrite: false });
        remoteValidation = await runOfficeHelper({ action: "validate", inputPath: verifyPath, kind });
        if (selectedBackend === "openxml" && remoteValidation.package?.fingerprint !== editResult.validation?.package?.fingerprint) {
          verificationIncomplete = true;
          toolCallContext.getStore()?.localWarnings?.push({ operation: "Office post-commit fingerprint verification", error: "Uploaded package fingerprint differs from the locally validated edit." });
        }
      } catch (error) {
        verificationIncomplete = true;
        toolCallContext.getStore()?.localWarnings?.push({ operation: "Office post-commit validation", error: safeToolErrorMessage(error) });
      }
    }
    await writeMutationAudit(toolName, {
      status: "success",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem),
      after: itemAuditSummary(afterRaw),
      officeEdit: { kind, backend: selectedBackend, changeCount: editResult.changeCount, operationTypes: args.operations.map((entry) => entry.type) },
      backupCreated: Boolean(backup),
      verificationIncomplete
    });
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(afterRaw),
      backend: selectedBackend,
      changes: editResult.changes,
      changeCount: editResult.changeCount,
      semanticDiff: officeSemanticDiff(kind, editResult.changes),
      validation: editResult.validation,
      remoteValidation,
      verificationIncomplete,
      backup,
      uploaded
    };
  } catch (error) {
    if (args.dryRun === false && args.confirmed === true) {
      await writeMutationAudit(toolName, {
        status: "failed",
        target: itemAuditSummary(rawItem),
        officeEdit: { kind, backend: selectedBackend, operationTypes: (args.operations || []).map((entry) => entry.type) },
        backupCreated: Boolean(backup),
        error: safeErrorInfo(error)
      });
    }
    throw error;
  } finally {
    await rm(transactionRoot, { recursive: true, force: true });
  }
}

function officeToolForKind(kind) {
  return `onedrive_${kind === "powerpoint" ? "powerpoint" : kind}_batch_update`;
}

async function officeBatchTransform(args = {}) {
  const toolName = "onedrive_office_batch_transform";
  const requested = args.items || [];
  if (args.dryRun === false && args.confirmed !== true) {
    return { dryRun: false, confirmed: false, requiredToUpdate: "Set dryRun: false and confirmed: true after reviewing the complete cross-file Office preview." };
  }
  if (args.dryRun === false) {
    const missingIdentity = requested.findIndex((item) => !hasExpectedIdentity(item));
    if (missingIdentity >= 0) return { dryRun: false, confirmed: true, requiredToUpdate: `Item ${missingIdentity} requires expectedName or expectedId.` };
  }

  const preflight = [];
  for (const [index, item] of requested.entries()) {
    const individualTool = officeToolForKind(item.kind);
    try {
      const preview = await officeBatchUpdate({
        ...item,
        dryRun: true,
        confirmed: false,
        createBackup: args.createBackup !== false,
        verify: args.verify !== false
      }, item.kind, individualTool);
      if (preview.noChanges) throw new Error(preview.requiredToUpdate || "No requested edits matched.");
      preflight.push({ index, request: item, individualTool, preview });
    } catch (error) {
      return {
        dryRun: args.dryRun !== false,
        confirmed: args.confirmed === true,
        preflightComplete: false,
        mutationStarted: false,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        items: preflight.map((entry) => entry.preview)
      };
    }
  }

  const proof = {
    items: preflight.map((entry) => ({
      id: entry.preview.item?.id,
      name: entry.preview.item?.name,
      eTag: entry.preview.item?.eTag,
      cTag: entry.preview.item?.cTag,
      kind: entry.request.kind,
      backend: entry.preview.backend,
      operations: entry.request.operations,
      changeCount: entry.preview.changeCount,
      outputFingerprint: entry.preview.validation?.package?.fingerprint
    }))
  };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    preflightComplete: true,
    mutationStarted: false,
    itemCount: preflight.length,
    totalChangeCount: preflight.reduce((sum, entry) => sum + Number(entry.preview.changeCount || 0), 0),
    items: preflight.map((entry) => ({ index: entry.index, kind: entry.request.kind, ...entry.preview }))
  };
  if (args.dryRun !== false) return previewWithToken(preview, toolName, proof);
  const tokenRequired = previewTokenRequiredResult(preview, toolName, proof, args.previewToken, "requiredToUpdate");
  if (tokenRequired) return tokenRequired;

  const completed = [];
  for (const entry of preflight) {
    try {
      const result = await officeBatchUpdate({
        ...entry.request,
        dryRun: false,
        confirmed: true,
        previewToken: entry.preview.previewToken,
        createBackup: args.createBackup !== false,
        verify: args.verify !== false
      }, entry.request.kind, entry.individualTool);
      completed.push({ index: entry.index, kind: entry.request.kind, result });
    } catch (error) {
      const remaining = preflight.slice(entry.index + 1).map((candidate) => ({ index: candidate.index, item: candidate.preview.item, kind: candidate.request.kind }));
      await writeMutationAudit(toolName, {
        status: "failed",
        officeBatch: { itemCount: preflight.length, completedCount: completed.length, failedIndex: entry.index },
        error: safeErrorInfo(error)
      });
      return {
        dryRun: false,
        confirmed: true,
        preflightComplete: true,
        mutationStarted: true,
        partialState: completed.length > 0,
        completed,
        failed: { index: entry.index, item: entry.preview.item, kind: entry.request.kind, error: safeToolErrorMessage(error) },
        remaining,
        recovery: completed.map((candidate) => ({ index: candidate.index, item: candidate.result.item, backup: candidate.result.backup }))
      };
    }
  }
  await writeMutationAudit(toolName, {
    status: "success",
    officeBatch: { itemCount: preflight.length, completedCount: completed.length, totalChangeCount: preview.totalChangeCount }
  });
  return {
    dryRun: false,
    confirmed: true,
    preflightComplete: true,
    mutationStarted: true,
    partialState: false,
    itemCount: completed.length,
    totalChangeCount: preview.totalChangeCount,
    completed
  };
}

function assertOfficeKind(info, kindName) {
  const kind = officeKinds[kindName];
  if (!kind) throw new Error(`Unknown Office helper kind: ${kindName}`);
  if (info.folder) throw new Error(`Item is a folder, not a ${kind.label} file: ${info.name}`);
  const extension = extname(info.name || "").toLowerCase();
  const mimeType = (info.file?.mimeType || "").toLowerCase();
  if (!kind.extensions.has(extension) && !kind.mimeTypes.has(mimeType)) {
    throw new Error(`Expected a ${kind.label}-compatible file. Got ${info.name || "unnamed item"} (${mimeType || "unknown MIME type"}).`);
  }
}

async function downloadOffice(args = {}, kindName) {
  const info = await getInfo(args);
  assertOfficeKind(info, kindName);
  return await downloadResolvedItem(info, {
    ...args,
    localPath: args.localPath || join(downloadRoot, kindName, info.name || `${kindName}-download`),
    overwrite: args.overwrite,
    _allowAlternateLocalPath: !args.localPath
  });
}

function exportFileName(name = "document", extension = ".pdf") {
  const base = basename(name, extname(name)) || "document";
  return `${base}${extension}`;
}

async function downloadExport(args = {}, formatName) {
  const exportFormats = {
    pdf: { graphFormat: "pdf", extension: ".pdf", label: "PDF" },
    text: { graphFormat: "text", extension: ".txt", label: "plain text" }
  };
  const format = exportFormats[formatName];
  if (!format) throw new Error(`Unknown export format: ${formatName}`);
  if (args.localPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.localPath), "Export", args);
  const info = await getInfo(args);
  if (info.folder) throw new Error(`Item is a folder, not an exportable document: ${info.name}`);
  const preferredTarget = args.localPath
    ? resolve(args.localPath)
    : join(downloadRoot, "export", exportFileName(info.name, format.extension));
  await assertNotLocalOneDriveSyncPathForWrite(preferredTarget, "Export", args);
  const reservation = await reserveLocalDestination(preferredTarget, {
    overwrite: args.overwrite === true,
    allowAlternate: !args.localPath
  });
  const target = reservation.path;
  const params = new URLSearchParams();
  params.set("format", format.graphFormat);
  try {
    const downloaded = await graphDownloadToFile(`${contentPath(args)}?${params.toString()}`, target, { reserved: reservation.reserved });
    return {
      item: info,
      localPath: target,
      bytesWritten: downloaded.bytesWritten,
      exportFormat: formatName,
      note: `Exported using Microsoft Graph format=${format.graphFormat}. Some file types may not support ${format.label} conversion.`
    };
  } catch (error) {
    if (reservation.reserved) await rm(target, { force: true });
    throw error;
  }
}

function truncateUtf8(text, maxBytes) {
  const buffer = Buffer.from(String(text), "utf8");
  if (buffer.length <= maxBytes) return { text: String(text), truncated: false, bytes: buffer.length };
  return {
    text: buffer.subarray(0, maxBytes).toString("utf8").replace(/\uFFFD$/u, ""),
    truncated: true,
    bytes: buffer.length
  };
}

async function preview(args = {}) {
  const info = await getInfo(args);
  const maxBytes = clampInteger(args.maxBytes, 65536, 1, 1048576);
  if (info.folder) return { item: info, preview: null, note: "Item is a folder; preview is only available for files." };

  if (args.preferExportText !== false && !isLikelyTextItem(info, args)) {
    try {
      const params = new URLSearchParams();
      params.set("format", "text");
      const exported = await graphLimitedBuffer(`${contentPath(args)}?${params.toString()}`, maxBytes);
      const previewText = exported.buffer.toString("utf8").replace(/\uFFFD$/u, "");
      return { item: info, preview: previewText, bytes: Buffer.byteLength(previewText, "utf8"), bytesRead: exported.bytesRead, truncated: exported.truncated, source: "graph-text-export" };
    } catch (error) {
      return { item: info, preview: null, source: "metadata", exportError: safeToolErrorMessage(error), note: "Graph text export was not available for this file." };
    }
  }

  assertTextReadable(info, args);
  const limited = await graphLimitedBuffer(contentPath(args), maxBytes);
  if (args.force !== true) assertNoBinaryNulls(limited.buffer);
  const previewText = limited.buffer.toString("utf8").replace(/\uFFFD$/u, "");
  return { item: info, preview: previewText, bytes: Buffer.byteLength(previewText, "utf8"), bytesRead: limited.bytesRead, truncated: limited.truncated, source: "text-read" };
}

function updateManifestPath(localPath, manifestPath) {
  return manifestPath ? resolve(manifestPath) : `${resolve(localPath)}.onedrive-update.json`;
}

async function updateFile(args = {}) {
  const remote = assertSafeRemotePath(args.remotePath, "remotePath");
  if (!remote) throw new Error("remotePath is required.");

  if (args.mode === "checkout") {
    if (args.localPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.localPath), "Checkout", args);
    if (args.manifestPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.manifestPath), "Checkout manifest", args);
    const info = await getInfo(args.itemId ? { itemId: args.itemId } : { path: remote });
    if (info.folder) throw new Error(`Cannot checkout a folder: ${info.name}`);
    const localPath = args.localPath ? resolve(args.localPath) : join(updateRoot, info.name || basename(remote));
    await assertNotLocalOneDriveSyncPathForWrite(localPath, "Checkout", args);
    const manifestPath = updateManifestPath(localPath, args.manifestPath);
    await assertNotLocalOneDriveSyncPathForWrite(manifestPath, "Checkout manifest", args);
    const manifestReservation = await reserveLocalDestination(manifestPath, {
      overwrite: args.overwriteManifest === true,
      allowAlternate: false
    });
    try {
      const downloaded = await download({
        ...(args.itemId ? { itemId: args.itemId } : { path: remote }),
        localPath,
        overwrite: args.overwriteLocal === true,
        allowLocalOneDriveSyncPath: args.allowLocalOneDriveSyncPath
      });
      const manifest = {
        version: 1,
        checkedOutAt: new Date().toISOString(),
        remotePath: remote,
        item: downloaded.item,
        localPath: downloaded.localPath
      };
      await writePrivateFileAtomic(manifestPath, JSON.stringify(manifest));
      return { mode: "checkout", ...downloaded, manifestPath };
    } catch (error) {
      if (manifestReservation.reserved) await rm(manifestPath, { force: true });
      throw error;
    }
  }

  if (args.mode !== "commit") throw new Error("mode must be checkout or commit.");
  const localPath = resolve(args.localPath || join(updateRoot, basename(remote)));
  await assertNotLocalOneDriveSyncPathForRead(localPath, "Commit", args);
  const manifestPath = updateManifestPath(localPath, args.manifestPath);
  let manifest = null;
  try {
    manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  } catch (error) {
    if (args.force !== true) throw new Error(`Could not read checkout manifest ${manifestPath}. Pass force: true only when intentionally overwriting without checkout metadata.`);
  }
  if (manifest && args.force !== true) {
    const manifestProblems = [];
    if (manifest.version !== 1) manifestProblems.push("version");
    if (manifest.remotePath && cleanPath(manifest.remotePath) !== cleanPath(remote)) manifestProblems.push("remotePath");
    if (manifest.localPath && resolve(manifest.localPath) !== localPath) manifestProblems.push("localPath");
    if (manifestPath && updateManifestPath(localPath, manifestPath) !== manifestPath) manifestProblems.push("manifestPath");
    if (manifestProblems.length) {
      throw new Error(`Checkout manifest does not match this commit request (${manifestProblems.join(", ")}). Re-checkout or pass force: true if you intend to override.`);
    }
  }

  const current = await getRawInfo({ path: remote });
  if (args.conflictCheck !== false && manifest?.item && args.force !== true) {
    const checked = manifest.item;
    const changed = [
      checked.id && current.id !== checked.id ? "id" : null,
      checked.eTag && current.eTag !== checked.eTag ? "eTag" : null,
      checked.cTag && current.cTag !== checked.cTag ? "cTag" : null,
      Number.isFinite(checked.size) && current.size !== checked.size ? "size" : null,
      checked.lastModifiedDateTime && current.lastModifiedDateTime !== checked.lastModifiedDateTime ? "lastModifiedDateTime" : null
    ].filter(Boolean);
    if (changed.length) {
      throw new Error(`Remote file changed since checkout (${changed.join(", ")}). Re-checkout or pass force: true if you intend to overwrite.`);
    }
  }

  let backup = null;
  if (args.createBackup !== false) {
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const backupPath = join(backupRoot, `${stamp}-${randomUUID()}-${basename(remote)}`);
    backup = await download({ path: remote, localPath: backupPath, overwrite: false });
  }

  const uploaded = await upload({
    localPath,
    remotePath: remote,
    conflictBehavior: "replace",
    allowLocalOneDriveSyncPath: args.allowLocalOneDriveSyncPath,
    ifMatch: args.conflictCheck !== false && args.force !== true ? manifest?.item?.eTag : undefined,
    auditTool: "onedrive_update_file"
  });
  const verified = args.verify !== false
    ? await bestEffortLocalWrite("post-commit remote verification", async () => await getInfo({ path: remote }))
    : null;
  if (manifest && args.force !== true) {
    const updatedManifest = {
      ...manifest,
      committedAt: new Date().toISOString(),
      remotePath: remote,
      localPath,
      item: verified || uploaded.item
    };
    await bestEffortLocalWrite("update manifest write", async () => await writePrivateFile(manifestPath, JSON.stringify(updatedManifest, null, 2)));
  }
  return {
    mode: "commit",
    remotePath: remote,
    localPath,
    manifestPath,
    backup,
    uploaded,
    verified,
    verificationIncomplete: args.verify !== false && !verified,
    note: "Committed local edits after checkout-manifest conflict checks."
  };
}

async function recent(args = {}) {
  const params = new URLSearchParams();
  params.set("$top", String(clampInteger(args.limit, 50, 1, 200)));
  const result = await graph(`/me/drive/recent?${params.toString()}`);
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(result.value || []));
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), count: (result.value || []).length };
}

async function largeFiles(args = {}) {
  const minBytes = args.minBytes ?? 104857600;
  const matches = [];
  const result = await scan({
    ...args,
    includeFiles: true,
    includeFolders: false,
    maxResults: 1,
    format: "full",
    onItem: (item) => {
      if (item.file && (item.size || 0) >= minBytes) matches.push(simplifyItem(item));
    }
  });
  const files = matches
    .sort((left, right) => (right.size || 0) - (left.size || 0))
    .slice(0, clampInteger(args.limit, 50, 1, 200))
    .map((item) => formatSimplifiedItem(item, args.format));
  return {
    filters: { minBytes },
    scanSummary: result.summary,
    count: files.length,
    itemsReturned: files.length,
    scanItemsReturned: result.summary.returned,
    items: files,
    note: "count/itemsReturned report files matching the large-file filter; scanSummary.returned reports the bounded traversal sample."
  };
}

function duplicateKey(item) {
  const hashes = item.file?.hashes || {};
  const hash = hashes.quickXorHash || hashes.sha1Hash || hashes.sha256Hash;
  if (hash) return `hash:${hash}`;
  return `name-size:${String(item.name || "").toLowerCase()}:${item.size || 0}`;
}

async function duplicates(args = {}) {
  const groups = new Map();
  const result = await scan({
    ...args,
    includeFiles: true,
    includeFolders: false,
    maxResults: 1,
    format: "full",
    onItem: (item) => {
      if (!item.file) return;
      const simplified = simplifyItem(item);
      const key = duplicateKey(simplified);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(simplified);
    }
  });
  const duplicates = [...groups.entries()]
    .filter(([, items]) => items.length > 1)
    .map(([key, items]) => ({
      key,
      count: items.length,
      totalBytes: items.reduce((sum, item) => sum + (item.size || 0), 0),
      items: items.map((item) => formatSimplifiedItem(item, args.format))
    }))
    .sort((left, right) => right.totalBytes - left.totalBytes)
    .slice(0, clampInteger(args.limit, 50, 1, 200));
  return {
    scanSummary: result.summary,
    duplicateGroups: duplicates.length,
    groupsReturned: duplicates.length,
    scanItemsReturned: result.summary.returned,
    groups: duplicates,
    note: "duplicateGroups/groupsReturned report duplicate groups after scanning; scanSummary.returned reports the bounded traversal sample."
  };
}

async function sharingAudit(args = {}, publicOnly = false) {
  const scanResult = await scan({
    ...args,
    includeFiles: true,
    includeFolders: args.includeFolders !== false,
    maxResults: Math.min(clampInteger(args.maxItems, 1000, 1, 50000), 5000),
    format: "full"
  });
  const matches = [];
  const candidates = scanResult.items || [];
  const limit = clampInteger(args.limit, 50, 1, 200);
  let auditedCount = 0;
  let errorCount = 0;
  let matchingItemCount = 0;
  const errors = [];
  for (let offset = 0; offset < candidates.length && matches.length < limit; offset += 20) {
    const chunk = candidates.slice(offset, offset + 20);
    let batch;
    try {
      batch = await batchPermissions({
        items: chunk.map((item) => ({ itemId: item.id })),
        format: "compact"
      });
    } catch (error) {
      errorCount += chunk.length;
      if (errors.length < 10) errors.push({ error: safeToolErrorMessage(error), itemIds: chunk.map((item) => item.id) });
      continue;
    }
    for (const [index, result] of (batch.items || []).entries()) {
      const item = chunk[index];
      if (!item) continue;
      if (result.error) {
        errorCount += 1;
        if (errors.length < 10) errors.push({ itemId: item.id, status: result.status, error: result.error });
        continue;
      }
      auditedCount += 1;
      const permissions = publicOnly
        ? result.permissions.filter((permission) => permission.permissionKind === "anonymous_link")
        : result.permissions.filter((permission) => args.includeOwnerPermissions === true ? !permission.inheritedFrom : isExplicitSharingPermission(permission));
      if (permissions.length) {
        matchingItemCount += 1;
        if (matches.length < limit) {
          matches.push({ item: formatSimplifiedItem(item, "compact"), permissions, count: permissions.length });
        }
      }
    }
  }
  const unauditedCount = Math.max(0, candidates.length - auditedCount - errorCount);
  const unseenCandidateCount = Math.max(0, (scanResult.summary?.matched || candidates.length) - candidates.length);
  const scanResultTruncated = Boolean(scanResult.summary?.resultTruncated);
  const resultTruncated = matchingItemCount > matches.length || unauditedCount > 0 || scanResultTruncated;
  const incomplete = Boolean(scanResult.summary?.traversalTruncated || scanResultTruncated || unseenCandidateCount || errorCount || resultTruncated);
  return {
    scanSummary: scanResult.summary,
    count: matches.length,
    itemsReturned: matches.length,
    scanItemsReturned: scanResult.summary.returned,
    candidateCount: candidates.length,
    auditedCount,
    errorCount,
    unauditedCount,
    unseenCandidateCount,
    matchingItemCount,
    resultTruncated,
    incomplete,
    errors,
    items: matches,
    note: incomplete
      ? "Sharing audit is incomplete because traversal was bounded, permission reads failed, or the result limit was reached. Review the audit counters and errors."
      : publicOnly
      ? "Returned items with anonymous sharing links."
      : (args.includeOwnerPermissions === true
        ? "Returned items with explicit sharing permissions plus owner grants."
        : "Returned items with explicit non-owner sharing permissions.")
  };
}

async function upload(args = {}) {
  const localPath = resolve(args.localPath);
  await assertNotLocalOneDriveSyncPathForRead(localPath, "Upload", args);
  const destinationPath = args.itemId ? (args.remotePath || `item:${args.itemId}`) : remotePath(args);
  const fileStat = await stat(localPath);
  if (!fileStat.isFile()) throw new Error(`Not a file: ${localPath}`);
  const uploadMode = args.uploadMode || "auto";
  const auditTool = args.auditTool || "onedrive_upload";
  try {
    let response;
    if (fileStat.size > 0 && (uploadMode === "session" || (uploadMode === "auto" && fileStat.size > simpleUploadLimit))) {
      const sessionTarget = args.itemId
        ? { endpoints: [`/me/drive/items/${encodeURIComponent(args.itemId)}/createUploadSession`], name: basename(destinationPath) }
        : undefined;
      response = await uploadLarge({ ...args, localPath, remotePath: destinationPath, sessionTarget }, fileStat);
    } else {
      if (uploadMode === "simple" && fileStat.size > simpleUploadLimit) {
        throw new Error(`Simple upload only supports files up to ${simpleUploadLimit} bytes. Use uploadMode: "session" or "auto".`);
      }
      const stream = createReadStream(localPath);
      const targetPath = args.itemId
        ? `/me/drive/items/${encodeURIComponent(args.itemId)}/content`
        : uploadPath(destinationPath, args.conflictBehavior || "fail");
      const result = await graph(targetPath, {
        method: "PUT",
        body: stream,
        duplex: "half",
        headers: {
          "Content-Type": "application/octet-stream",
          ...(args.ifMatch ? { "If-Match": args.ifMatch } : {})
        }
      });
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
      response = { item: simplifyItem(result), localPath, bytesUploaded: fileStat.size, uploadMode: "simple" };
    }
    if (args.skipAudit !== true) {
      await writeMutationAudit(auditTool, {
        status: "success",
        target: { remotePath: destinationPath },
        after: itemAuditSummary(response.item),
        localPath,
        bytes: fileStat.size
      });
    }
    return response;
  } catch (error) {
    if (args.skipAudit !== true) {
      await writeMutationAudit(auditTool, {
        status: "failed",
        target: { remotePath: destinationPath },
        localPath,
        bytes: fileStat.size,
        error: safeErrorInfo(error)
      });
    }
    throw error;
  }
}

function normalizeChunkSize(value = defaultUploadChunkSize) {
  const chunkSize = Number(value);
  if (!Number.isInteger(chunkSize) || chunkSize < uploadChunkUnit || chunkSize > maxUploadChunkSize) {
    throw new Error(`chunkSize must be an integer between ${uploadChunkUnit} and ${maxUploadChunkSize} bytes.`);
  }
  if (chunkSize % uploadChunkUnit !== 0) {
    throw new Error(`chunkSize must be a multiple of ${uploadChunkUnit} bytes.`);
  }
  return chunkSize;
}

async function uploadLarge(args = {}, fileStat) {
  const chunkSize = normalizeChunkSize(args.chunkSize ?? defaultUploadChunkSize);
  const sessionTarget = args.sessionTarget || await uploadSessionTarget(args.remotePath);
  const session = await createUploadSession(sessionTarget, args.conflictBehavior || "fail", args.ifMatch);
  if (!session.uploadUrl) throw new Error("Microsoft Graph did not return an uploadUrl.");

  let position = 0;
  let uploaded = 0;
  let finalItem = null;
  const handle = await open(args.localPath, "r");
  try {
    while (position < fileStat.size) {
      const length = Math.min(chunkSize, fileStat.size - position);
      const buffer = Buffer.allocUnsafe(length);
      const { bytesRead } = await handle.read(buffer, 0, length, position);
      if (bytesRead <= 0) throw new Error(`Could not read upload chunk at byte ${position}.`);
      const body = bytesRead === buffer.length ? buffer : buffer.subarray(0, bytesRead);
      const end = position + bytesRead - 1;
      let response;
      try {
        response = await graph(assertTrustedUploadSessionUrl(session.uploadUrl), {
          method: "PUT",
          skipAuth: true,
          body,
          headers: {
            "Content-Range": `bytes ${position}-${end}/${fileStat.size}`
          },
          maxRetries: 4
        });
      } catch (error) {
        throw new Error(`Upload session failed for byte range ${position}-${end}/${fileStat.size}: ${safeToolErrorMessage(error)}`);
      }
      uploaded = end + 1;
      if (response?.id) finalItem = response;
      if (!response?.id && Array.isArray(response?.nextExpectedRanges)) {
        const nextStart = Number(String(response.nextExpectedRanges[0] || "").split("-")[0]);
        if (Number.isFinite(nextStart) && nextStart !== uploaded) {
          throw new Error(`Upload session expected next byte ${nextStart}, but local upload position is ${uploaded}.`);
        }
      }
      position += bytesRead;
    }
  } finally {
    await handle.close();
  }

  if (!finalItem) throw new Error("Upload session completed local chunks but Microsoft Graph did not return a final drive item.");
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([finalItem]));
  return {
    item: simplifyItem(finalItem),
    localPath: args.localPath,
    bytesUploaded: uploaded,
    uploadMode: "session",
    chunkSize
  };
}

async function createUploadSession(sessionTarget, conflictBehavior, ifMatch) {
  const bodies = [
    {
      item: {
        "@odata.type": "microsoft.graph.driveItemUploadableProperties",
        "@microsoft.graph.conflictBehavior": conflictBehavior,
        name: sessionTarget.name
      }
    },
    {
      item: {
        "@microsoft.graph.conflictBehavior": conflictBehavior,
        name: sessionTarget.name
      }
    },
    null
  ];
  const errors = [];
  for (const endpoint of sessionTarget.endpoints) {
    for (const body of bodies) {
      try {
        return await graph(endpoint, {
          method: "POST",
          ...(ifMatch ? { headers: { "If-Match": ifMatch } } : {}),
          ...(body ? { body: JSON.stringify(body) } : {})
        });
      } catch (error) {
        errors.push(`${endpoint}: ${safeToolErrorMessage(error)}`);
      }
    }
  }
  throw new Error(`Could not create upload session. Tried ${errors.length} compatible request shapes. Last error: ${errors.at(-1)}`);
}

async function writeText(args = {}) {
  const destinationPath = remotePath(args);
  try {
    const result = await graph(uploadPath(destinationPath, args.conflictBehavior || "fail"), {
      method: "PUT",
      body: Buffer.from(args.content, "utf8"),
      headers: { "Content-Type": "text/plain; charset=utf-8" }
    });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
    const response = { item: simplifyItem(result), bytesUploaded: Buffer.byteLength(args.content, "utf8") };
    await writeMutationAudit("onedrive_write_text", {
      status: "success",
      target: { remotePath: destinationPath },
      after: itemAuditSummary(response.item),
      bytes: response.bytesUploaded
    });
    return response;
  } catch (error) {
    await writeMutationAudit("onedrive_write_text", {
      status: "failed",
      target: { remotePath: destinationPath },
      bytes: Buffer.byteLength(args.content || "", "utf8"),
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function createFolder(args = {}) {
  const name = assertSafeItemName(args.name, "name");
  assertAtMostOneSelector(args, "folder parent", [
    { label: "parentItemId", keys: ["parentItemId"] },
    { label: "parentPath", keys: ["parentPath"] },
    { label: "parentPreset", keys: ["parentPreset"] }
  ]);
  const endpoint = args.parentItemId
    ? `/me/drive/items/${encodeURIComponent(args.parentItemId)}/children`
    : childrenPath({
        path: args.parentPath || "",
        preset: args.parentPreset,
        relativePath: args.parentRelativePath
      });
  try {
    const result = await graph(endpoint, {
      method: "POST",
      body: JSON.stringify({
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": args.conflictBehavior || "fail"
      })
    });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
    const item = simplifyItem(result);
    await writeMutationAudit("onedrive_create_folder", {
      status: "success",
      target: { parentPath: args.parentPath, parentItemId: args.parentItemId, name },
      after: itemAuditSummary(item)
    });
    return item;
  } catch (error) {
    await writeMutationAudit("onedrive_create_folder", {
      status: "failed",
      target: { parentPath: args.parentPath, parentItemId: args.parentItemId, name },
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function rename(args = {}) {
  const newName = assertSafeItemName(args.newName, "newName");
  requireNonRootTarget(args, "Rename");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Rename refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Rename");
  const item = simplifyItem(current);
  if (args.dryRun !== false) {
    return { dryRun: true, wouldRename: item, newName };
  }
  if (args.confirmed !== true) {
    return {
      dryRun: false,
      confirmed: false,
      wouldRename: item,
      newName,
      requiredToRename: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      dryRun: false,
      confirmed: true,
      wouldRename: item,
      newName,
      requiredToRename: "Provide expectedName or expectedId for live renames."
    };
  }
  try {
    const result = await graph(itemMutationBase(current), {
      method: "PATCH",
      headers: mutationMatchHeaders(current),
      body: JSON.stringify({ name: newName })
    });
    await cacheMovedOrRenamedItem(current, result);
    const renamed = simplifyItem(result);
    await writeMutationAudit("onedrive_rename", {
      status: "success",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      after: itemAuditSummary(renamed)
    });
    return { dryRun: false, confirmed: true, renamed };
  } catch (error) {
    await writeMutationAudit("onedrive_rename", {
      status: "failed",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      requestedName: newName,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function moveItem(args = {}) {
  if (args.newName) assertSafeItemName(args.newName, "newName");
  requireNonRootTarget(args, "Move");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Move refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Move");
  const parentReference = await resolveDestinationParent(args);
  const body = { parentReference: { id: parentReference.id } };
  if (args.newName) body.name = args.newName;
  const item = simplifyItem(current);
  if (args.dryRun !== false) {
    return { dryRun: true, wouldMove: item, destination: parentReference, newName: args.newName || null };
  }
  if (args.confirmed !== true) {
    return {
      dryRun: false,
      confirmed: false,
      wouldMove: item,
      destination: parentReference,
      newName: args.newName || null,
      requiredToMove: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      dryRun: false,
      confirmed: true,
      wouldMove: item,
      destination: parentReference,
      newName: args.newName || null,
      requiredToMove: "Provide expectedName or expectedId for live moves."
    };
  }
  try {
    const result = await graph(itemMutationBase(current), {
      method: "PATCH",
      headers: mutationMatchHeaders(current),
      body: JSON.stringify(body)
    });
    await cacheMovedOrRenamedItem(current, result);
    const moved = simplifyItem(result);
    await writeMutationAudit("onedrive_move", {
      status: "success",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      after: itemAuditSummary(moved),
      destination: parentReference
    });
    return { dryRun: false, confirmed: true, moved };
  } catch (error) {
    await writeMutationAudit("onedrive_move", {
      status: "failed",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      destination: parentReference,
      newName: args.newName || null,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function pollCopyMonitor(monitorUrl, timeoutSeconds = 60) {
  const trustedMonitorUrl = assertTrustedCopyMonitorUrl(monitorUrl);
  const safeMonitorUrl = safeDisplayPath(trustedMonitorUrl);
  const deadline = Date.now() + timeoutSeconds * 1000;
  let last = null;
  while (Date.now() < deadline) {
    const response = await graph(trustedMonitorUrl, { skipAuth: true, returnResponse: true, maxRetries: 3, redirect: "manual" });
    last = response.body && !(response.body instanceof ArrayBuffer) ? response.body : null;
    if (response.status === 303) {
      return {
        complete: true,
        status: response.status,
        resourceLocation: response.headers.get("location") ? safeDisplayPath(response.headers.get("location")) : null,
        monitorUrl: safeMonitorUrl
      };
    }
    if (response.ok && last?.status && !["notStarted", "running", "inProgress"].includes(String(last.status))) {
      return { complete: true, status: response.status, monitorUrl: safeMonitorUrl, monitor: sanitizeAuditValue(last) };
    }
    await sleep(2000);
  }
  return { complete: false, timeoutSeconds, monitorUrl: safeMonitorUrl, monitor: sanitizeAuditValue(last) };
}

async function copyItem(args = {}) {
  if (args.newName) assertSafeItemName(args.newName, "newName");
  requireNonRootTarget(args, "Copy");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Copy refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Copy");
  const parentReference = await resolveDestinationParent(args);
  const item = simplifyItem(current);
  if (args.dryRun !== false) {
    return { dryRun: true, wouldCopy: item, destination: parentReference, newName: args.newName || null };
  }
  if (args.confirmed !== true) {
    return {
      dryRun: false,
      confirmed: false,
      wouldCopy: item,
      destination: parentReference,
      newName: args.newName || null,
      requiredToCopy: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      dryRun: false,
      confirmed: true,
      wouldCopy: item,
      destination: parentReference,
      newName: args.newName || null,
      requiredToCopy: "Provide expectedName or expectedId for live copies."
    };
  }
  try {
    const response = await graph(`${itemMutationBase(current)}/copy`, {
      method: "POST",
      returnResponse: true,
      body: JSON.stringify({
        parentReference: { id: parentReference.id },
        ...(args.newName ? { name: args.newName } : {})
      })
    });
    if (!response.ok) throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Copy failed" });
    const monitorUrl = response.headers.get("location");
    const result = {
      dryRun: false,
      confirmed: true,
      accepted: response.status === 202 || response.ok,
      status: response.status,
      source: item,
      monitorUrl: monitorUrl ? safeDisplayPath(monitorUrl) : null
    };
    if (args.waitForCompletion && monitorUrl) {
      try {
        result.monitor = await pollCopyMonitor(monitorUrl, args.timeoutSeconds ?? 60);
      } catch (error) {
        result.monitorError = safeToolErrorMessage(error);
      }
    }
    await writeMutationAudit("onedrive_copy", {
      status: result.monitorError ? "accepted-monitor-failed" : "success",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      destination: parentReference,
      newName: args.newName || null,
      monitorError: result.monitorError,
      graphRequestId: response.graphRequestId
    });
    return result;
  } catch (error) {
    await writeMutationAudit("onedrive_copy", {
      status: "failed",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      destination: parentReference,
      newName: args.newName || null,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function createSharingLink(args = {}) {
  requireNonRootTarget(args, "Create sharing link");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Create sharing link refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Create sharing link");
  const includePermissionDiff = args.includePermissionDiff !== false;
  const beforePermissions = includePermissionDiff ? await permissionList({ itemId: current.id }, "compact") : null;
  const warnings = [];
  const requestedScope = args.scope || "anonymous";
  if (requestedScope === "organization") {
    try {
      const driveInfo = await graph("/me/drive");
      if (driveInfo.driveType === "personal") {
        warnings.push("Organization-scoped links are not meaningful on personal OneDrive drives; Microsoft Graph may reject or reinterpret this scope.");
      }
    } catch {
      warnings.push("Could not verify drive type before previewing organization-scoped link creation.");
    }
  }
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldCreate: {
      item: simplifyItem(current),
      type: args.type || "view",
      scope: requestedScope,
      passwordProvided: Boolean(args.password),
      expirationDateTime: args.expirationDateTime,
      retainInheritedPermissions: args.retainInheritedPermissions
    },
    warnings,
    ...(includePermissionDiff ? {
      beforePermissions,
      beforePermissionCount: beforePermissions.length
    } : {})
  };
  const previewProof = {
    item: { id: current.id, name: current.name },
    type: args.type || "view",
    scope: requestedScope,
    password: args.password,
    expirationDateTime: args.expirationDateTime,
    retainInheritedPermissions: args.retainInheritedPermissions
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_create_sharing_link", previewProof),
      requiredToCreate: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToCreate: "Provide expectedName or expectedId for live sharing-link creation."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_create_sharing_link", previewProof, args.previewToken, "requiredToCreate");
  if (previewTokenRequired) return previewTokenRequired;
  const body = {
    type: args.type || "view",
    scope: args.scope || "anonymous"
  };
  if (args.password) body.password = args.password;
  if (args.expirationDateTime) body.expirationDateTime = args.expirationDateTime;
  if (typeof args.retainInheritedPermissions === "boolean") {
    body.retainInheritedPermissions = args.retainInheritedPermissions;
  }
  try {
    const result = await graph(`${itemMutationBase(current)}/createLink`, {
      method: "POST",
      body: JSON.stringify(body)
    });
    const afterPermissions = includePermissionDiff
      ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: current.id }, "compact"))
      : null;
    const permissionDiff = includePermissionDiff && afterPermissions ? diffPermissions(beforePermissions, afterPermissions) : null;
    await writeMutationAudit("onedrive_create_sharing_link", {
      status: "success",
      target: itemAuditSummary(current),
      link: {
        type: body.type,
        scope: body.scope,
        passwordProvided: Boolean(args.password),
        expirationDateTime: args.expirationDateTime,
        retainInheritedPermissions: body.retainInheritedPermissions
      },
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      afterPermissions: includePermissionDiff && afterPermissions ? afterPermissions.map(permissionAuditSummary) : undefined,
      permissionDiff: includePermissionDiff && permissionDiff ? permissionDiffAuditSummary(permissionDiff) : undefined
    });
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(current),
      permission: result,
      verificationIncomplete: includePermissionDiff && !afterPermissions,
      ...(includePermissionDiff ? {
        beforePermissions,
        afterPermissions,
        permissionDiff
      } : {})
    };
  } catch (error) {
    await writeMutationAudit("onedrive_create_sharing_link", {
      status: "failed",
      target: itemAuditSummary(current),
      link: {
        type: body.type,
        scope: body.scope,
        passwordProvided: Boolean(args.password),
        expirationDateTime: args.expirationDateTime,
        retainInheritedPermissions: body.retainInheritedPermissions
      },
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function normalizeInviteRecipients(recipients = []) {
  return recipients.map((recipient, index) => {
    const keys = ["email", "alias", "objectId"].filter((key) => recipient[key]);
    if (keys.length !== 1) {
      throw new Error(`Invite recipient at index ${index} must include exactly one of email, alias, or objectId.`);
    }
    return { [keys[0]]: recipient[keys[0]] };
  });
}

function recipientKinds(recipients = []) {
  return recipients.map((recipient) => ["email", "alias", "objectId"].find((key) => recipient[key]) || "unknown");
}

function inviteBody(args = {}) {
  const body = {
    recipients: normalizeInviteRecipients(args.recipients || []),
    roles: [args.role || "read"],
    sendInvitation: args.sendInvitation === true,
    requireSignIn: args.requireSignIn !== false
  };
  if (args.message) body.message = args.message;
  if (args.password) body.password = args.password;
  if (args.expirationDateTime) body.expirationDateTime = args.expirationDateTime;
  if (typeof args.retainInheritedPermissions === "boolean") {
    body.retainInheritedPermissions = args.retainInheritedPermissions;
  }
  return body;
}

function inviteAuditSummary(invite = {}) {
  return {
    recipientCount: invite.recipientCount,
    recipientKinds: invite.recipientKinds,
    role: invite.role,
    sendInvitation: invite.sendInvitation,
    requireSignIn: invite.requireSignIn,
    messageProvided: invite.messageProvided,
    passwordProvided: invite.passwordProvided,
    expirationDateTime: invite.expirationDateTime,
    retainInheritedPermissions: invite.retainInheritedPermissions
  };
}

async function invitePermission(args = {}) {
  requireNonRootTarget(args, "Invite permission");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Invite permission refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Invite permission");
  const body = inviteBody(args);
  const includePermissionDiff = args.includePermissionDiff !== false;
  const beforePermissions = includePermissionDiff ? await permissionList({ itemId: current.id }, "compact") : null;
  const safeInvite = {
    item: simplifyItem(current),
    recipientCount: body.recipients.length,
    recipientKinds: recipientKinds(body.recipients),
    role: body.roles[0],
    sendInvitation: body.sendInvitation,
    requireSignIn: body.requireSignIn,
    messageProvided: Boolean(args.message),
    passwordProvided: Boolean(args.password),
    expirationDateTime: args.expirationDateTime,
    retainInheritedPermissions: args.retainInheritedPermissions
  };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldInvite: safeInvite,
    ...(includePermissionDiff ? {
      beforePermissions,
      beforePermissionCount: beforePermissions.length
    } : {})
  };
  const previewProof = {
    item: { id: current.id, name: current.name },
    recipients: body.recipients,
    role: body.roles[0],
    sendInvitation: body.sendInvitation,
    requireSignIn: body.requireSignIn,
    message: args.message,
    password: args.password,
    expirationDateTime: args.expirationDateTime,
    retainInheritedPermissions: args.retainInheritedPermissions
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_invite_permission", previewProof),
      requiredToInvite: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToInvite: "Provide expectedName or expectedId for live permission invitation."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_invite_permission", previewProof, args.previewToken, "requiredToInvite");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const result = await graph(`${itemMutationBase(current)}/invite`, {
      method: "POST",
      body: JSON.stringify(body)
    });
    const afterPermissions = includePermissionDiff
      ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: current.id }, "compact"))
      : null;
    const permissionDiff = includePermissionDiff && afterPermissions ? diffPermissions(beforePermissions, afterPermissions) : null;
    await writeMutationAudit("onedrive_invite_permission", {
      status: "success",
      target: itemAuditSummary(current),
      invite: inviteAuditSummary(safeInvite),
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      afterPermissions: includePermissionDiff && afterPermissions ? afterPermissions.map(permissionAuditSummary) : undefined,
      permissionDiff: includePermissionDiff && permissionDiff ? permissionDiffAuditSummary(permissionDiff) : undefined
    });
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(current),
      invite: safeInvite,
      permissions: Array.isArray(result.value) ? result.value.map((permission) => simplifyPermission(permission, args.format)) : result,
      verificationIncomplete: includePermissionDiff && !afterPermissions,
      ...(includePermissionDiff ? {
        beforePermissions,
        afterPermissions,
        permissionDiff
      } : {})
    };
  } catch (error) {
    await writeMutationAudit("onedrive_invite_permission", {
      status: "failed",
      target: itemAuditSummary(current),
      invite: inviteAuditSummary(safeInvite),
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function assertPermissionPresent(permissions = [], permissionId) {
  if (!permissions.some((permission) => permission.id === permissionId)) {
    throw new Error(`Permission ${permissionId} was not found on the target item. Refusing to continue.`);
  }
}

function findPermissionById(permissions = [], permissionId) {
  return permissions.find((permission) => permission.id === permissionId) || null;
}

async function preflightRevokePermission(args = {}, options = {}) {
  requireNonRootTarget(args, "Revoke permission");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Revoke permission refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Revoke permission");
  const includePermissions = options.includePermissions !== false;
  const permissionsForPreflight = await permissionList({ itemId: current.id }, "compact");
  assertPermissionPresent(permissionsForPreflight, args.permissionId);
  const permission = findPermissionById(permissionsForPreflight, args.permissionId);
  const revocable = isRevocablePermission(permission);
  const warnings = revocable ? [] : [`Permission ${args.permissionId} is a ${permission?.permissionKind || "non-revocable"} permission and should not be revoked with this tool.`];
  return {
    targetArgs: args,
    rawItem: current,
    item: simplifyItem(current),
    permission,
    revocable,
    warnings,
    beforePermissions: includePermissions ? permissionsForPreflight : [],
    includePermissions
  };
}

async function revokePermission(args = {}) {
  const preflight = await preflightRevokePermission(args, { includePermissions: args.includePermissions });
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldRevoke: {
      item: preflight.item,
      permissionId: args.permissionId,
      permission: preflight.permission,
      revocable: preflight.revocable
    },
    warnings: preflight.warnings,
    ...(preflight.includePermissions ? {
      beforePermissions: preflight.beforePermissions,
      beforePermissionCount: preflight.beforePermissions.length
    } : {})
  };
  const previewProof = {
    item: { id: preflight.rawItem.id, name: preflight.rawItem.name },
    permissionId: args.permissionId
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_revoke_permission", previewProof),
      requiredToRevoke: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToRevoke: "Provide expectedName or expectedId for live permission revocation."
    };
  }
  if (!preflight.revocable) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToRevoke: "Choose a revocable sharing permission. Owner and inherited permissions are refused."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_revoke_permission", previewProof, args.previewToken, "requiredToRevoke");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const response = await graph(`${itemIdBase(preflight.rawItem.id)}/permissions/${encodeURIComponent(args.permissionId)}`, {
      method: "DELETE",
      returnResponse: true
    });
    if (!response.ok) throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Revoke permission failed" });
    const afterPermissions = preflight.includePermissions
      ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: preflight.rawItem.id }, "compact"))
      : null;
    const permissionDiff = preflight.includePermissions && afterPermissions ? diffPermissions(preflight.beforePermissions, afterPermissions) : null;
    await writeMutationAudit("onedrive_revoke_permission", {
      status: "success",
      target: itemAuditSummary(preflight.rawItem),
      permissionId: args.permissionId,
      beforePermissions: preflight.includePermissions ? preflight.beforePermissions.map(permissionAuditSummary) : undefined,
      afterPermissions: preflight.includePermissions && afterPermissions ? afterPermissions.map(permissionAuditSummary) : undefined,
      permissionDiff: preflight.includePermissions && permissionDiff ? permissionDiffAuditSummary(permissionDiff) : undefined,
      graphRequestId: response.graphRequestId
    });
    return {
      dryRun: false,
      confirmed: true,
      item: preflight.item,
      permissionId: args.permissionId,
      verificationIncomplete: preflight.includePermissions && !afterPermissions,
      ...(preflight.includePermissions ? {
        beforePermissions: preflight.beforePermissions,
        afterPermissions,
        permissionDiff
      } : {})
    };
  } catch (error) {
    await writeMutationAudit("onedrive_revoke_permission", {
      status: "failed",
      target: itemAuditSummary(preflight.rawItem),
      permissionId: args.permissionId,
      beforePermissions: preflight.includePermissions ? preflight.beforePermissions.map(permissionAuditSummary) : undefined,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function batchRevokePermissions(args = {}) {
  const items = args.items || [];
  const warnings = batchMutationWarnings();
  if (args.dryRun === false) {
    if (args.confirmed !== true) {
      return {
        dryRun: false,
        confirmed: false,
        count: items.length,
        warnings,
        requiredToRevoke: "Set dryRun: false and confirmed: true after explicit user confirmation."
      };
    }
    const missingExpected = items.filter((item) => !hasExpectedIdentity(item));
    if (missingExpected.length) {
      return {
        dryRun: false,
        confirmed: true,
        count: items.length,
        warnings,
        requiredToRevoke: "Provide expectedName or expectedId for every item in a live batch permission revoke.",
        missingExpectedCount: missingExpected.length
      };
    }
  }

  const preflight = [];
  const preflightErrors = [];
  for (const [index, item] of items.entries()) {
    try {
      preflight.push(await preflightRevokePermission(item, { includePermissions: args.includePermissions }));
    } catch (error) {
      preflightErrors.push({ index, target: item, error: safeToolErrorMessage(error) });
    }
  }
  if (preflightErrors.length) {
    return {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: preflightErrors,
      requiredToRevoke: "Fix every preflight error before running a batch permission revoke."
    };
  }
  const nonRevocable = preflight
    .map((entry, index) => entry.revocable ? null : { index, target: entry.targetArgs, permission: entry.permission, warnings: entry.warnings })
    .filter(Boolean);
  if (args.dryRun === false && nonRevocable.length) {
    return {
      dryRun: false,
      confirmed: args.confirmed === true,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: nonRevocable.map((entry) => ({
        index: entry.index,
        target: entry.target,
        error: entry.warnings.join(" ")
      })),
      requiredToRevoke: "Choose only revocable sharing permissions. Owner and inherited permissions are refused."
    };
  }
  const previewProof = {
    items: preflight.map((entry) => ({
      id: entry.rawItem.id,
      name: entry.rawItem.name,
      permissionId: entry.targetArgs.permissionId
    })),
    operation: "batch-revoke"
  };

  if (args.dryRun !== false) {
    return {
      dryRun: true,
      confirmed: args.confirmed === true,
      count: preflight.length,
      warnings,
      ...issuePreviewToken("onedrive_batch_revoke_permissions", previewProof),
      results: preflight.map((entry) => ({
        wouldRevoke: { item: entry.item, permissionId: entry.targetArgs.permissionId, permission: entry.permission, revocable: entry.revocable },
        warnings: entry.warnings,
        ...(entry.includePermissions ? {
          beforePermissions: entry.beforePermissions,
          beforePermissionCount: entry.beforePermissions.length
        } : {})
      }))
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(
    { dryRun: false, confirmed: true, count: preflight.length, warnings },
    "onedrive_batch_revoke_permissions",
    previewProof,
    args.previewToken,
    "requiredToRevoke"
  );
  if (previewTokenRequired) return previewTokenRequired;

  const results = [];
  for (const [index, entry] of preflight.entries()) {
    try {
      const response = await graph(`${itemIdBase(entry.rawItem.id)}/permissions/${encodeURIComponent(entry.targetArgs.permissionId)}`, {
        method: "DELETE",
        returnResponse: true
      });
      if (!response.ok) throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Batch revoke permission failed" });
      const afterPermissions = entry.includePermissions
        ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: entry.rawItem.id }, "compact"))
        : null;
      const permissionDiff = entry.includePermissions && afterPermissions ? diffPermissions(entry.beforePermissions, afterPermissions) : null;
      results.push({
        item: entry.item,
        permissionId: entry.targetArgs.permissionId,
        verificationIncomplete: entry.includePermissions && !afterPermissions,
        ...(entry.includePermissions ? { beforePermissions: entry.beforePermissions, afterPermissions, permissionDiff } : {})
      });
    } catch (error) {
      await writeMutationAudit("onedrive_batch_revoke_permissions", {
        status: "failed",
        targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
        partialResults: results.map((result) => ({ item: itemAuditSummary(result.item), permissionId: result.permissionId })),
        failedIndex: index,
        error: safeErrorInfo(error)
      });
      return {
        dryRun: false,
        confirmed: true,
        failed: true,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        count: preflight.length,
        warnings,
        partialResults: results
      };
    }
  }
  await writeMutationAudit("onedrive_batch_revoke_permissions", {
    status: "success",
    targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
    results: results.map((result) => ({
      item: itemAuditSummary(result.item),
      permissionId: result.permissionId,
      permissionDiff: result.permissionDiff ? permissionDiffAuditSummary(result.permissionDiff) : undefined
    }))
  });
  return { dryRun: false, confirmed: true, count: results.length, warnings, results };
}

async function preflightMoveItem(item = {}, destination) {
  if (item.newName) assertSafeItemName(item.newName, "newName");
  requireNonRootTarget(item, "Move");
  const current = await getRawInfo(item);
  if (current.root) throw new Error("Move refuses to operate on the OneDrive root.");
  assertExpectedItem(current, item, "Move");
  return { targetArgs: item, rawItem: current, item: simplifyItem(current), destination };
}

async function batchMove(args = {}) {
  const items = args.items || [];
  const warnings = batchMutationWarnings();
  if (args.dryRun === false) {
    if (args.confirmed !== true) {
      return {
        dryRun: false,
        confirmed: false,
        count: items.length,
        warnings,
        requiredToMove: "Set dryRun: false and confirmed: true after explicit user confirmation."
      };
    }
    const missingExpected = items.filter((item) => !hasExpectedIdentity(item));
    if (missingExpected.length) {
      return {
        dryRun: false,
        confirmed: true,
        count: items.length,
        warnings,
        requiredToMove: "Provide expectedName or expectedId for every item in a live batch move.",
        missingExpectedCount: missingExpected.length
      };
    }
  }

  const destination = await resolveDestinationParent(args);
  const preflight = [];
  const preflightErrors = [];
  for (const [index, item] of items.entries()) {
    try {
      preflight.push(await preflightMoveItem(item, destination));
    } catch (error) {
      preflightErrors.push({ index, target: item, error: safeToolErrorMessage(error) });
    }
  }
  if (preflightErrors.length) {
    return {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      destination,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: preflightErrors,
      requiredToMove: "Fix every preflight error before running a batch move."
    };
  }

  if (args.dryRun !== false) {
    return {
      dryRun: true,
      confirmed: args.confirmed === true,
      destination,
      count: preflight.length,
      warnings,
      results: preflight.map((entry) => ({
        wouldMove: entry.item,
        destination,
        newName: entry.targetArgs.newName || null
      }))
    };
  }

  const results = [];
  for (const [index, entry] of preflight.entries()) {
    try {
      const body = { parentReference: { id: destination.id } };
      if (entry.targetArgs.newName) body.name = entry.targetArgs.newName;
      const result = await graph(itemIdBase(entry.rawItem.id), {
        method: "PATCH",
        headers: mutationMatchHeaders(entry.rawItem),
        body: JSON.stringify(body)
      });
      await cacheMovedOrRenamedItem(entry.rawItem, result);
      results.push({ before: entry.item, moved: simplifyItem(result), newName: entry.targetArgs.newName || null });
    } catch (error) {
      await writeMutationAudit("onedrive_batch_move", {
        status: "failed",
        destination,
        targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
        partialResults: results.map((result) => ({ before: itemAuditSummary(result.before), after: itemAuditSummary(result.moved) })),
        failedIndex: index,
        error: safeErrorInfo(error)
      });
      return {
        dryRun: false,
        confirmed: true,
        failed: true,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        destination,
        count: preflight.length,
        warnings,
        partialResults: results
      };
    }
  }
  await writeMutationAudit("onedrive_batch_move", {
    status: "success",
    destination,
    targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
    results: results.map((result) => ({
      before: itemAuditSummary(result.before),
      after: itemAuditSummary(result.moved)
    }))
  });
  return { dryRun: false, confirmed: true, destination, count: results.length, warnings, results };
}

async function deleteItem(args = {}) {
  requireNonRootTarget(args, "Delete");
  const rawItem = await getRawInfo(args);
  if (rawItem.root) throw new Error("Delete refuses to operate on the OneDrive root.");
  assertExpectedItem(rawItem, args, "Delete");
  const item = simplifyItem(rawItem);
  const previewProof = { items: [{ id: rawItem.id, name: rawItem.name }], operation: "delete" };
  if (args.dryRun !== false) {
    return previewWithToken({ dryRun: true, wouldDelete: item }, "onedrive_delete", previewProof);
  }
  if (args.confirmed !== true) {
    return {
      dryRun: false,
      confirmed: false,
      wouldDelete: item,
      requiredToDelete: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!args.expectedName && !args.expectedId) {
    return {
      dryRun: false,
      confirmed: true,
      wouldDelete: item,
      requiredToDelete: "Provide expectedName or expectedId for live deletes."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(
    { dryRun: false, confirmed: true, wouldDelete: item },
    "onedrive_delete",
    previewProof,
    args.previewToken,
    "requiredToDelete"
  );
  if (previewTokenRequired) return previewTokenRequired;
  try {
    await graph(itemMutationBase(rawItem), { method: "DELETE", headers: mutationMatchHeaders(rawItem) });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([{ ...rawItem, deleted: {} }]));
    await writeMutationAudit("onedrive_delete", {
      status: "success",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem)
    });
    return { dryRun: false, confirmed: true, deleted: item };
  } catch (error) {
    await writeMutationAudit("onedrive_delete", {
      status: "failed",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem),
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function validateRestoreArgs(args = {}) {
  if (args.newName) assertSafeItemName(args.newName, "newName");
  if (args.destinationParentPath) assertSafeRemotePath(args.destinationParentPath, "destinationParentPath");
  if (args.destinationParentRelativePath) assertSafeRemotePath(args.destinationParentRelativePath, "destinationParentRelativePath");
  if (args.destinationParentPreset) {
    resolvePresetPath(args, {
      pathField: "destinationParentPath",
      presetField: "destinationParentPreset",
      relativeField: "destinationParentRelativePath"
    });
  }
}

async function restoreDeleted(args = {}) {
  validateRestoreArgs(args);
  if (args.expectedId && args.expectedId !== args.itemId) {
    throw new Error(`Restore expected item ID ${args.expectedId}, but got ${args.itemId}. Refusing to continue.`);
  }
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldRestore: {
      itemId: args.itemId,
      destinationParentPath: args.destinationParentPath,
      destinationParentItemId: args.destinationParentItemId,
      destinationParentPreset: args.destinationParentPreset,
      destinationParentRelativePath: args.destinationParentRelativePath,
      newName: args.newName
    },
    permissionNote: "OneDrive Personal restore may require Microsoft Graph delegated Files.ReadWrite.All in addition to the plugin's standard Files.ReadWrite scope."
  };
  const previewProof = {
    itemId: args.itemId,
    destinationParentPath: args.destinationParentPath,
    destinationParentItemId: args.destinationParentItemId,
    destinationParentPreset: args.destinationParentPreset,
    destinationParentRelativePath: args.destinationParentRelativePath,
    newName: args.newName
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_restore_deleted", previewProof),
      requiredToRestore: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!args.expectedId) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToRestore: "Provide expectedId for live restores."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_restore_deleted", previewProof, args.previewToken, "requiredToRestore");
  if (previewTokenRequired) return previewTokenRequired;
  const body = {};
  if (args.destinationParentPath || args.destinationParentItemId || args.destinationParentPreset) {
    const parent = await resolveDestinationParent(args);
    body.parentReference = { id: parent.id };
  }
  if (args.newName) body.name = args.newName;
  try {
    const result = await graph(`/me/drive/items/${encodeURIComponent(args.itemId)}/restore`, {
      method: "POST",
      body: JSON.stringify(body)
    });
    const restored = simplifyItem(result);
    await writeMutationAudit("onedrive_restore_deleted", {
      status: "success",
      target: { itemId: args.itemId },
      after: itemAuditSummary(restored)
    });
    return { dryRun: false, confirmed: true, restored };
  } catch (error) {
    await writeMutationAudit("onedrive_restore_deleted", {
      status: "failed",
      target: { itemId: args.itemId },
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function textResult(value, isError = false) {
  const localWarnings = toolCallContext.getStore()?.localWarnings || [];
  const payload = localWarnings.length
    ? (value && typeof value === "object" && !Array.isArray(value)
        ? { ...value, localWarnings }
        : { message: value, localWarnings })
    : value;
  const text = typeof payload === "string" ? payload : JSON.stringify(payload, null, 2);
  return { content: [{ type: "text", text }], isError };
}

function sendResult(id, result) {
  process.stdout.write(JSON.stringify({ jsonrpc: "2.0", id, result }) + "\n");
}

function sendError(id, code, message) {
  process.stdout.write(JSON.stringify({ jsonrpc: "2.0", id, error: { code, message } }) + "\n");
}

async function callTool(name, args = {}) {
  switch (name) {
    case "onedrive_config": {
      const status = publicConfig();
      if (args.checkToken) {
        try {
          await getAccessToken();
          status.accessTokenAvailable = true;
        } catch (error) {
          status.accessTokenAvailable = false;
          status.tokenCheckError = safeToolErrorMessage(error);
        }
      }
      return textResult(status);
    }
    case "onedrive_auth_device_start":
      return textResult(await startDeviceLogin(args));
    case "onedrive_auth_device_poll":
      return textResult(await pollDeviceLogin(args));
    case "onedrive_logout": {
      authGeneration += 1;
      deviceLoginGeneration += 1;
      tokenCache = null;
      pendingDevice = null;
      tokenRefreshPromise = null;
      if (args.deleteKeychainToken === true && args.confirmed !== true) {
        return textResult({
          memoryCleared: true,
          keychainTokenDeleted: false,
          confirmed: false,
          requiredToDelete: "Set confirmed: true after explicit user confirmation to delete the OneDrive Keychain refresh token."
        });
      }
      return textResult({ memoryCleared: true, keychainTokenDeleted: args.deleteKeychainToken ? deleteKeychainToken() : false });
    }
    case "onedrive_me":
      return textResult(await graph("/me"));
    case "onedrive_drive":
      return textResult(await graph("/me/drive"));
    case "onedrive_doctor":
      return textResult(await doctor(args));
    case "onedrive_presets":
      return textResult({ pathPresets: pathPresets(), configPath });
    case "onedrive_list":
      return textResult(await list(args));
    case "onedrive_list_all":
      return textResult(await listAll(args));
    case "onedrive_scan":
      return textResult(await scan(args));
    case "onedrive_search":
      return textResult(await search(args));
    case "onedrive_search_all":
      return textResult(await searchAll(args));
    case "onedrive_find":
      return textResult(await find(args));
    case "onedrive_find_all":
      return textResult(await findAll(args));
    case "onedrive_delta":
      return textResult(await delta(args));
    case "onedrive_sync_status":
      return textResult(await syncStatus(args));
    case "onedrive_cache_refresh":
      return textResult(await cacheRefresh(args));
    case "onedrive_cache_clear":
      return textResult(await clearMetadataCache());
    case "onedrive_content_index_refresh":
      return textResult(await contentIndexRefresh(args));
    case "onedrive_content_search":
      return textResult(await contentSearch(args));
    case "onedrive_office_index_refresh":
      return textResult(await officeIndexRefresh(args));
    case "onedrive_office_search":
      return textResult(await officeContentSearch(args));
    case "onedrive_content_index_clear":
      return textResult(await clearContentIndex());
    case "onedrive_office_capabilities":
      return textResult(await officeCapabilities());
    case "onedrive_office_validate":
      return textResult(await inspectRemoteOfficePackage(args, args.expectedKind || null, "validate"));
    case "onedrive_word_get_document":
      return textResult(await inspectRemoteOfficePackage(args, "word"));
    case "onedrive_excel_get_workbook":
      return textResult(await inspectRemoteOfficePackage(args, "excel"));
    case "onedrive_powerpoint_get_presentation":
      return textResult(await inspectRemoteOfficePackage(args, "powerpoint"));
    case "onedrive_word_batch_update":
      return textResult(await officeBatchUpdate(args, "word", "onedrive_word_batch_update"));
    case "onedrive_excel_batch_update":
      return textResult(await officeBatchUpdate(args, "excel", "onedrive_excel_batch_update"));
    case "onedrive_powerpoint_batch_update":
      return textResult(await officeBatchUpdate(args, "powerpoint", "onedrive_powerpoint_batch_update"));
    case "onedrive_office_backups":
      return textResult(await officeBackups(args));
    case "onedrive_office_compare_backup":
      return textResult(await officeCompareBackup(args));
    case "onedrive_office_restore_backup":
      return textResult(await officeRestoreBackup(args));
    case "onedrive_office_batch_transform":
      return textResult(await officeBatchTransform(args));
    case "onedrive_get_info":
      return textResult(await getInfo(args));
    case "onedrive_read_text":
      return textResult(await readText(args));
    case "onedrive_preview":
      return textResult(await preview(args));
    case "onedrive_download":
      return textResult(await download(args));
    case "onedrive_download_excel":
      return textResult(await downloadOffice(args, "excel"));
    case "onedrive_download_word":
      return textResult(await downloadOffice(args, "word"));
    case "onedrive_download_powerpoint":
      return textResult(await downloadOffice(args, "powerpoint"));
    case "onedrive_export_pdf":
      return textResult(await downloadExport(args, "pdf"));
    case "onedrive_export_text":
      return textResult(await downloadExport(args, "text"));
    case "onedrive_upload":
      return textResult(await upload(args));
    case "onedrive_write_text":
      return textResult(await writeText(args));
    case "onedrive_create_folder":
      return textResult(await createFolder(args));
    case "onedrive_rename":
      return textResult(await rename(args));
    case "onedrive_move":
      return textResult(await moveItem(args));
    case "onedrive_copy":
      return textResult(await copyItem(args));
    case "onedrive_create_sharing_link":
      return textResult(await createSharingLink(args));
    case "onedrive_invite_permission":
      return textResult(await invitePermission(args));
    case "onedrive_revoke_permission":
      return textResult(await revokePermission(args));
    case "onedrive_batch_revoke_permissions":
      return textResult(await batchRevokePermissions(args));
    case "onedrive_permissions":
      return textResult(await permissions(args));
    case "onedrive_batch_get_info":
      return textResult(await batchGetInfo(args));
    case "onedrive_batch_permissions":
      return textResult(await batchPermissions(args));
    case "onedrive_batch_download":
      return textResult(await batchDownload(args));
    case "onedrive_batch_delete":
      return textResult(await batchDelete(args));
    case "onedrive_batch_move":
      return textResult(await batchMove(args));
    case "onedrive_update_file":
      return textResult(await updateFile(args));
    case "onedrive_recent":
      return textResult(await recent(args));
    case "onedrive_large_files":
      return textResult(await largeFiles(args));
    case "onedrive_duplicates":
      return textResult(await duplicates(args));
    case "onedrive_shared_by_me":
      return textResult(await sharingAudit(args, false));
    case "onedrive_public_links":
      return textResult(await sharingAudit(args, true));
    case "onedrive_restore_deleted":
      return textResult(await restoreDeleted(args));
    case "onedrive_audit_recent":
      return textResult(await auditRecent(args));
    case "onedrive_audit_export":
      return textResult(await auditExport(args));
    case "onedrive_audit_clear":
      return textResult(await auditClear(args));
    case "onedrive_delete":
      return textResult(await deleteItem(args));
    default:
      return textResult(`Unknown tool: ${name}`, true);
  }
}

async function handleRequest(message) {
  const { id, method, params = {} } = message;
  try {
    if (method === "initialize") {
      sendResult(id, {
        protocolVersion: "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: { name: "onedrive", version: pluginManifest.version || "0.1.0" }
      });
      return;
    }
    if (method === "tools/list") {
      sendResult(id, { tools });
      return;
    }
    if (method === "tools/call") {
      const validation = validateToolArguments(params.name, params.arguments || {});
      if (!validation.ok) {
        sendResult(id, textResult(validation.error, true));
        return;
      }
      const result = await toolCallContext.run({
        localWarnings: [],
        lastGraphRequestId: null,
        lastMutationGraphRequestId: null,
        metadataCacheWrites: 0
      }, async () => {
        try {
          return await callTool(params.name, validation.args || {});
        } catch (error) {
          return textResult(safeToolErrorMessage(error), true);
        }
      });
      sendResult(id, result);
      return;
    }
    if (method?.startsWith("notifications/")) return;
    if (id !== undefined) sendError(id, -32601, `Method not found: ${method}`);
  } catch (error) {
    if (id !== undefined) sendResult(id, textResult(safeToolErrorMessage(error), true));
    else console.error(error);
  }
}

let buffer = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
  buffer += chunk;
  for (;;) {
    const newline = buffer.indexOf("\n");
    if (newline === -1) break;
    const line = buffer.slice(0, newline).trim();
    buffer = buffer.slice(newline + 1);
    if (!line) continue;
    try {
      void handleRequest(JSON.parse(line));
    } catch (error) {
      sendError(null, -32700, `Parse error: ${error.message}`);
    }
  }
});
