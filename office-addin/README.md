# Optional Office companion

The Office companion is a local, manual-paste task-pane add-in for the document currently open in Word, Excel, or PowerPoint. Version `1.1.1` complements the headless OneDrive/Open XML tools; it does not provide a remote command channel or replace those tools.

## Security and protocol

- The task pane exposes `globalThis.CodexOfficeCompanion.getCapabilities()` and `executeCommand(value)`.
- Capability negotiation uses protocol `codex-office-companion/1`, reports the active host and exact requirement set, and lists only commands supported in that host.
- A command can be passed directly or in an envelope: `{"protocolVersion":"codex-office-companion/1","command":{...}}`.
- `executeCommand` limits the UTF-8 encoded, JSON-serialized input to 65,536 bytes. This also protects callers that bypass the task-pane textarea.
- Commands reject unknown properties and invalid types instead of coercing values. Only the active document is accessible; network callbacks and telemetry remain disabled.

## Exact command schemas

Properties described as optional must be omitted when unused. JSON property names and enum values are case-sensitive.

### Word (`WordApi 1.3`)

- `{"type":"replaceText","find":"old","replace":"new","all":true,"matchCase":false}`
  - `find` is a required non-empty string.
  - `replace` is an optional string and defaults to `""` (delete each match).
  - `all` is an optional boolean and defaults to `true`; `false` changes only the first match.
  - `matchCase` is an optional boolean and defaults to `false`.
- `{"type":"setSelectedText","text":"Reviewed copy"}`
  - `text` is required and must be a string. An empty string deletes the selection.
  - A non-empty text selection is required before execution.
- `{"type":"insertParagraph","text":"Next paragraph","location":"after"}`
  - `text` is required and must be a string. An empty string inserts a blank paragraph.
  - `location` is optional, accepts `"before"` or `"after"`, and defaults to `"after"`.

### Excel (`ExcelApi 1.7`)

Every Excel command requires non-empty string `sheet` and `address` properties. Matrix properties must be non-empty rectangular two-dimensional arrays whose row and column counts exactly match the target range.

- `{"type":"setRange","sheet":"Sheet1","address":"A1:B2","values":[[1,2],[3,4]]}`
  - At least one of `values`, `formulas`, or `numberFormat` is required.
  - `values` and `formulas` cells can be strings, booleans, finite numbers, or `null`.
  - `numberFormat` cells must be strings.
- `{"type":"clearRange","sheet":"Sheet1","address":"A1:B2","applyTo":"Contents"}`
  - `applyTo` is optional and defaults to `"Contents"`.
  - Accepted values are `"All"`, `"Contents"`, `"Formats"`, `"Hyperlinks"`, and `"RemoveHyperlinks"`.
- `{"type":"formatRange","sheet":"Sheet1","address":"A1:B2","fillColor":"#DDEBFF","bold":true}`
  - At least one formatting property is required.
  - Optional non-empty strings: `fillColor`, `fontColor`, and `fontName`.
  - Optional `fontSize` must be a finite number from 1 through 409.
  - Optional `bold` and `italic` values must be booleans.
  - Optional `numberFormat` is a string matrix matching the target range.

### PowerPoint (`PowerPointApi 1.5`)

- `{"type":"setSelectedText","text":"Reviewed copy"}`
  - `text` is required and must be a string. An empty string deletes the selection.
  - A non-empty text selection is required. The command edits the selected text range, not every character in the selected shape.
- `{"type":"setSelectedTextStyle","fontName":"Aptos","fontSize":20,"color":"#123456","bold":true,"italic":false,"underline":true}`
  - At least one style property is required, along with a non-empty text selection.
  - `fontName` and `color` must be non-empty strings; `fontSize` must be a finite number from 1 through 400.
  - `bold`, `italic`, and `underline` must be booleans. Underline maps to the Office `Single` or `None` underline enum.
- `{"type":"deleteSelectedShapes"}`
  - At least one shape must be selected. Only selected shapes are deleted.

The textarea is initialized with a valid starter command for the active host. Review and adjust its document-specific names, ranges, text, or selections before applying it.

## Ribbon launch surface

The manifest preserves the base task-pane entry for older Office clients and adds one `VersionOverridesV1_0` ribbon command for each supported host: `Document` (Word), `Workbook` (Excel), and `Presentation` (PowerPoint). In clients that support `AddinCommands 1.1`, open **Home → Codex OneDrive → Open Companion**. That button uses Office's built-in `ShowTaskpane` action to open the same local task pane; it does not use `ExecuteFunction` and does not add another command to the nine-command companion protocol.

The Home-tab command is an independent launch surface when the Mac **Insert → My Add-ins** flyout displays a sideloaded development add-in as disabled. The base flyout entry remains a compatibility fallback and is not removed.

## Run-scoped local HTTPS host and Mac sideload

The checked-in manifest loads only from `https://127.0.0.1:3443`. It references exact 32×32 and 64×64 base icons plus the required 16×16, 32×32, and 80×80 ribbon icons. `serve.mjs` binds only to that loopback address and serves only the companion web assets. Leave unrelated development servers, including anything on port 3000, untouched.

Quit Word, Excel, and PowerPoint, choose a unique run ID, and start the guarded host runner from the plugin root:

```sh
RUN_ID="office-$(date -u +%Y%m%dT%H%M%SZ)"
zsh office-addin/host-test-runner.zsh "$RUN_ID"
```

The runner refuses a pre-existing run directory, manifest, or port-3443 listener. It then:

- creates a one-day, run-scoped `CA:TRUE,pathlen:0` development root and a signed `CA:FALSE`, `serverAuth` loopback leaf, and records both SHA-256 hashes;
- trusts only that temporary root for SSL to `127.0.0.1` in the user's login Keychain, then verifies the leaf through the installed Keychain trust;
- generates a temporary manifest with a new UUID, run-specific display name, and run-specific filename;
- serves the plugin root on `https://127.0.0.1:3443` and verifies the page with the exact certificate;
- hard-links only the run-specific manifest into the Word, Excel, and PowerPoint `wef` directories using Microsoft's ID-prefixed developer-registration filename; and
- stays open so its exit trap can remove the exact manifests, stop only its recorded server PID, delete the exact temporary root and user trust settings by SHA-256 with `security delete-certificate -t -Z`, verify the loopback leaf is no longer trusted, remove only run-created empty `wef` directories, remove its private temporary directory, and verify port 3443 is closed.

The Keychain command may prompt for the current macOS account password. Keep the runner terminal open throughout the host matrix. It prints `OFFICE_HOST_SETUP_READY` with the run ID, PID, CA and leaf certificate hashes, manifest ID, and manifest name when setup is complete. Before cleanup it prints request counts for the task pane, script, and each icon without logging query strings, cookies, authorization values, or other request headers.

## Exact real-host matrix

Use a new unsaved scratch document in each host with AutoSave off. Open **Home → Codex OneDrive → Open Companion** for the run-specific **Codex OneDrive Office Companion Test …** add-in. If the ribbon command is unavailable on an older client, use **Insert → My Add-ins** as the base-manifest fallback. Record the returned operation/change count or the expected error substring for every row. Close each scratch document without saving.

Valid Word commands (`WordApi 1.3`):

| Setup | Command | Expected evidence |
| --- | --- | --- |
| Body contains `old text` | `{"type":"replaceText","find":"old text","replace":"new text","all":true,"matchCase":false}` | All matches change; `operation:"replaceText"` |
| Select disposable text | `{"type":"setSelectedText","text":"Reviewed copy"}` | Only the selection changes; `changed:1` |
| Place a disposable insertion point/selection | `{"type":"insertParagraph","text":"Next paragraph","location":"after"}` | One paragraph is inserted after it |

Valid Excel commands (`ExcelApi 1.7`):

| Setup | Command | Expected evidence |
| --- | --- | --- |
| Blank `Sheet1!A1:B2` | `{"type":"setRange","sheet":"Sheet1","address":"A1:B2","values":[[1,2],[3,4]],"numberFormat":[["0","0"],["0","0"]]}` | Exact 2×2 values/formats; `changed:1` |
| Populated `Sheet1!A1:B2` | `{"type":"clearRange","sheet":"Sheet1","address":"A1:B2","applyTo":"Contents"}` | Contents clear while formatting remains |
| Blank `Sheet1!A1:B1` | `{"type":"formatRange","sheet":"Sheet1","address":"A1:B1","fillColor":"#DDEBFF","fontColor":"#123456","fontName":"Aptos","fontSize":14,"bold":true,"italic":false,"numberFormat":[["0.00","0.00"]]}` | Exact range formatting; `changed:1` |

Valid PowerPoint commands (`PowerPointApi 1.5`):

| Setup | Command | Expected evidence |
| --- | --- | --- |
| Select text inside a disposable text box | `{"type":"setSelectedText","text":"Reviewed copy"}` | Only selected characters change; `changed:1` |
| Select text inside that box | `{"type":"setSelectedTextStyle","fontName":"Aptos","fontSize":20,"color":"#123456","bold":true,"italic":false,"underline":true}` | Exact style; underline is Single |
| Select one disposable shape | `{"type":"deleteSelectedShapes"}` | Only that shape is deleted; returned count matches selection |

Run the following rejection payloads in the named host. The document must remain unchanged:

| Host | Payload | Expected error substring |
| --- | --- | --- |
| Any | `{"protocolVersion":"wrong","command":{"type":"deleteSelectedShapes"}}` | `protocolVersion` |
| Any | `{"command":{"type":"deleteSelectedShapes"}}` | `protocolVersion` |
| Any | `[]` | `JSON object` |
| Any | `{"type":3}` | `type must be a string` |
| PowerPoint | `{"type":"deleteSelectedShapes","unknown":true}` | `Unknown command property` |
| Word | `{"type":"replaceText","find":""}` | `find must not be empty` |
| Word | `{"type":"replaceText","find":"x","replace":1}` | `replace must be a string` |
| Word | `{"type":"replaceText","find":"x","all":"yes"}` | `all must be a boolean` |
| Word | `{"type":"insertParagraph","text":"x","location":"start"}` | `location must be one of` |
| Word | `{"type":"setSelectedText","text":null}` | `text must be a string` |
| Word, with no text selected | `{"type":"setSelectedText","text":"x"}` | `non-empty Word text selection` |
| Excel | `{"type":"setRange","sheet":"Sheet1","address":"A1"}` | `requires values` |
| Excel | `{"type":"formatRange","sheet":"Sheet1","address":"A1"}` | `at least one formatting property` |
| Excel | `{"type":"setRange","sheet":"Sheet1","address":"A1:B2","values":[[1],[2,3]]}` | `rectangular` |
| Excel | `{"type":"setRange","sheet":"Sheet1","address":"A1:B2","values":[[1]]}` | `must match range dimensions` |
| Excel | `{"type":"clearRange","sheet":"Sheet1","address":"A1","applyTo":"ResetContents"}` | `applyTo must be one of` |
| Excel | `{"type":"formatRange","sheet":"Sheet1","address":"A1","bold":1}` | `bold must be a boolean` |
| Excel | `{"type":"formatRange","sheet":"Sheet1","address":"A1","numberFormat":[[1]]}` | `must be a string` |
| PowerPoint | `{"type":"setSelectedTextStyle"}` | `at least one formatting property` |
| PowerPoint | `{"type":"setSelectedTextStyle","underline":"Single"}` | `underline must be a boolean` |
| PowerPoint | `{"type":"setSelectedTextStyle","fontSize":0}` | `between 1 and 400` |
| PowerPoint | `{"type":"setSelectedText","text":9}` | `text must be a string` |
| PowerPoint, with no text selected | `{"type":"setSelectedTextStyle","bold":true}` | `non-empty PowerPoint text selection` |
| PowerPoint, with no shape selected | `{"type":"deleteSelectedShapes"}` | `at least one selected shape` |

The serialized-size and non-JSON-number guards are direct-API cases. In the task-pane Web Inspector, verify these exact calls reject without changing the document:

```js
CodexOfficeCompanion.executeCommand({type: "setSelectedText", text: "😀".repeat(20000)})
// rejects with: exceeds 65536 bytes

CodexOfficeCompanion.executeCommand({type: "setRange", sheet: "Sheet1", address: "A1", values: [[NaN]]})
// rejects with: finite number

CodexOfficeCompanion.executeCommand({type: "formatRange", sheet: "Sheet1", address: "A1", fontSize: Infinity})
// rejects with: finite number

const cyclic = {type: "setSelectedText", text: "hello"}; cyclic.self = cyclic;
CodexOfficeCompanion.executeCommand(cyclic)
// rejects with: JSON-serializable
```

After completing the matrix, quit all three Office apps and press Return in the runner terminal. Require the final `OFFICE_HOST_CLEANUP_OK` line. Restart each Office app and confirm the run-specific companion no longer appears. Never clear shared Office caches; doing so can affect unrelated add-ins and requires separate consent.

## Validation and deployment

Run the offline companion suite from the plugin root:

```sh
node --check office-addin/taskpane.js
node --check office-addin/manifest-contract.mjs
node --check office-addin/serve.mjs
node --check office-addin/prepare-test-manifest.mjs
node office-addin/prepare-test-manifest.mjs --self-check
zsh -n office-addin/host-test-runner.zsh
node office-addin/taskpane-test.mjs
xmllint --noout office-addin/manifest.xml
sips -g pixelWidth -g pixelHeight office-addin/icon-16.png office-addin/icon-32.png office-addin/icon-64.png office-addin/icon-80.png
```

The test negotiates all three hosts, exercises all nine commands and rejection paths with mocked Office runtimes, checks starter commands, verifies the exact three-host `ShowTaskpane` ribbon contract, and validates every advertised icon dimension offline. It cannot replace a real Word, Excel, and PowerPoint host run.

For centralized deployment, copy the manifest, replace every `https://127.0.0.1:3443` URL with one immutable HTTPS origin you control, increment the four-part manifest version, host the task pane script, HTML, and icons at that origin, and deploy through Microsoft 365 integrated apps. Keep the Content Security Policy unchanged unless an explicitly reviewed transport is added.
