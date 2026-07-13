# OneDrive 0.4.0 Release Gate Report

Decision: Pass
Date: 2026-07-13
Generated: 2026-07-13T21:47:14Z
Source commit: `a94eb39f3cb5d21cfed305574a1a7d06f52b61ae` with release-candidate working-tree changes
Plugin version: `0.4.0+codex.20260713202830`
Office companion version: `1.1.1`
Tool contract: 72 exact tool names

## Current outcome

The rebuilt source and immutable installed-cache live betas each passed with 91 passes, zero failures, eight explicitly blocked checks, the exact 72-tool contract, and verified isolated cleanup. The installed cache has 44-file byte, mode, type, and symlink-target parity while preserving every older cache.

All required offline, source-live, installed-live, Office-host, cleanup, and parity gates are complete. The remaining blocked coverage is explicitly environmental or safety constrained and is covered by mocks or dry runs where possible.

## Completed offline gates

- Node 20 and Node 26 syntax and utility self-checks
- Exact 72-tool MCP contract
- Three-host Office manifest and 16×16, 32×32, 64×64, and 80×80 icons
- Prepackage negative tests, including the final-report regression
- All 38 Open XML operations: Word 9, Excel 17, PowerPoint 12
- Office security corpus and genuine LibreOffice package reopening
- All nine Office companion commands
- Mock Microsoft Graph suite
- Read-only OneDrive doctor and tenant matrix
- Zero-candidate cleanup preview
- Whitespace validation

## Final live and packaging evidence

- Source run `codex-beta-20260713t162523z-source`: 91 pass, 0 fail, 8 blocked, 99 total; runtime 855,447 ms.
- Installed run `codex-beta-20260713t162523z-installed`: 91 pass, 0 fail, 8 blocked, 99 total; runtime 918,660 ms.
- Both exact test roots and both isolated local work areas were cleaned.
- Immutable cache installed at `$CODEX_HOME/plugins/cache/personal/onedrive/0.4.0+codex.20260713202830`; 44 packaged files matched and older caches were preserved.

## Explicitly blocked live coverage

- Business/work-school Microsoft Graph Excel is unavailable on the personal drive; all eight operation types are mock-tested.
- Organization-only sharing requires a work or school tenant and remains mock-tested.
- Forced device-code polling and Keychain credential deletion remain blocked so the existing credential and consent are untouched.
- Live recycle-bin restore remains blocked; mock and dry-run coverage passed.
