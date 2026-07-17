# OneDrive 0.5.0 Release Gate Report

Decision: Beta Pending — Installed Cache Verification Required
Date: 2026-07-17
Generated: 2026-07-17T03:38:35Z
Tested source commit: `fd1b102b50bf0b0e8104c1b7c85848167e7d4bed`
Plugin version: `0.5.0+codex.20260717033601`
Tool contract: 84 exact tool names

## Current outcome

Release candidate `0.5.0+codex.20260717033601` passes the 84-tool contract, 153-case mock Graph suite, focused security and Office tests, exact source/cache parity, and six-check read-only live harnesses from both source and the new immutable installed cache. Full CRUD on the fresh connector process remains pending because the current task continues using the previously loaded cache.

## Explicitly constrained coverage

- Business Graph Excel and organization-only sharing are mock-tested because the live account is personal.
- Existing credentials, consent, and Keychain data remain untouched; forced login polling and Keychain deletion are excluded.
- Live mutation coverage used one isolated real-user beta folder on the previously installed build; the folder and all descendants were deleted after verification.
- Shared Office caches will not be cleared.
- Live recycle-bin restore remains excluded; native version restore has independent coverage.
- The newly installed cache cannot replace the MCP server already loaded by this task; a fresh task is required for the final guarded `write_text` live-replacement retest.
