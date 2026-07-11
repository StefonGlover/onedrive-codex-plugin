# Optional Office companion

This task-pane add-in applies a small, typed command set to the document currently open in Word, Excel, or PowerPoint. It complements the headless OneDrive/Open XML tools; it does not replace them.

Security properties:

- No remote command or telemetry endpoint is configured.
- The Content Security Policy blocks network connections from the task pane.
- Commands are pasted and reviewed locally, reject unknown properties, and operate only on the active document.
- Runtime requirement sets are checked before host-specific APIs are used.

At startup the companion exposes `globalThis.CodexOfficeCompanion` with `getCapabilities()` and `executeCommand()`. Capability responses use protocol `codex-office-companion/1`, name the active host and platform, report the exact requirement set, list only executable commands, and explicitly report that remote commands and telemetry are disabled. Commands may be pasted directly or wrapped in a negotiated envelope:

```json
{"protocolVersion":"codex-office-companion/1","command":{"type":"setSelectedText","text":"Reviewed copy"}}
```

Supported companion commands:

- Word (`WordApi 1.3`): `replaceText`, `setSelectedText`, `insertParagraph`
- Excel (`ExcelApi 1.7`): `setRange`, `clearRange`, `formatRange`
- PowerPoint (`PowerPointApi 1.5`): `setSelectedText`, `setSelectedTextStyle`, `deleteSelectedShapes`

Commands are capped at 64 KiB and reject unknown properties. Selection-oriented commands intentionally require a user-visible active selection, making the scope reviewable before execution.

To sideload during development, serve this directory on `https://localhost:3000/office-addin/` with a trusted local development certificate, then sideload `manifest.xml` using the platform-specific Office Add-ins development workflow. This optional component requires an active Office host; normal OneDrive plugin editing remains headless.

For production deployment, copy `manifest.xml`, replace every `https://localhost:3000` URL with one immutable HTTPS origin you control, increment the four-part manifest version, host `taskpane.html`, `taskpane.js`, and the referenced icons at that origin, then deploy the manifest through Microsoft 365 integrated apps/centralized deployment. Keep the task pane CSP unchanged unless an explicitly reviewed transport is added. The current companion has no MCP callback or remote execution endpoint by design.

Run `node office-addin/taskpane-test.mjs` before packaging. The test negotiates all three hosts with mocked Office runtimes and exercises the typed selection/range commands without launching Office.
