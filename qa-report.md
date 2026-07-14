# OneDrive 0.5.0 Release Gate Report

Decision: Finalizing Immutable Build
Date: 2026-07-14
Generated: 2026-07-14T05:01:00Z
Tested source commit: `ad711b5b08bc81d289bbcb8ca008d5053603f651`
Plugin version: `0.5.0+codex.20260714050128`
Tool contract: 84 exact tool names

## Current outcome

Release `0.5.0+codex.20260714050128` is ready for immutable-cache installation. The exact 84-tool contract, 140-case mock Graph suite, 78 headless Office operations, semantic anchors, structured text patches, managed workspaces, watches, Office security corpus, genuine-package reopening, cleanup, and packaging checks pass. The source and installed live betas are explicitly recorded as waived by user direction, so this report does not label them as passed.

## Explicitly constrained coverage

- Business Graph Excel and organization-only sharing are mock-tested because the live account is personal.
- Existing credentials, consent, and Keychain data remain untouched; forced login polling and Keychain deletion are excluded.
- Shared Office caches will not be cleared.
- Live recycle-bin restore remains excluded; native version restore has independent coverage.
- The user explicitly waived the remaining source and installed live beta runs and directed Codex to wrap up immediately.
