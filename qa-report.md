# OneDrive 0.5.0 Release Gate Report

Decision: Pass
Date: 2026-07-19
Generated: 2026-07-19T18:30:35Z
Tested source commit: `e0a2da72a301221239f786cd19ecbe737d4874db`
Plugin version: `0.5.0+codex.20260719183035`
Tool contract: 84 exact tool names

## Current outcome

The current source and immutable installed build each passed the full 106-check real-user beta with 96 passes, zero failures, ten explicit environment or safety blocks, exact coverage of all 84 tools, verified remote cleanup, and no local residue. The 153-check mock Graph suite, encrypted-auth tests, Open XML/security/reopening gates, package validation, and source/cache byte-mode-type-symlink parity also pass. The same build is live through an auto-restarting, outbound-only OpenAI tunnel on the Synology DS923+, and ChatGPT verified a healthy token plus read-only account and root access.

## Bugs and improvements validated

- Native version restore now uses Microsoft Graph's `restoreVersion` action and passes live plus mock verification.
- The beta harness uses one configurable Office Python runtime, cleans partial fixture setup, persists optional JSON reports, emits compact progress, and bounds child Graph request latency.
- Live search no longer multiplies transient retries across layers.
- Named-share testing rejects the owner's own mailbox and safely classifies a verified `sharingFailed` result only when no permission was created.
- Excel Open XML edits now preserve declared compatibility namespaces, reject undeclared `mc:Ignorable` prefixes, copy row formatting and height for appended table rows, and extend matching conditional-formatting and data-validation ranges.
- Excel `deleteTableRow` now compacts a bounded table row, preserves surrounding worksheet content, translates ordinary relative formulas, shrinks bounded single-column shared-formula groups without changing formulas or cached values, shifts native hyperlinks, shrinks matching ranges, preserves row formatting, and maintains Excel-required ascending cell order.
- Linux and NAS deployments use an AES-256-GCM encrypted token vault with a separately mounted owner-only key, atomic writes, symlink refusal, and strict permission validation.
- Synology startup now normalizes an uploaded encrypted token to `0600` before dropping privileges, fixing DSM/File Station's permissive upload mode without weakening vault checks.
- The DS923+ image pins the classic DSM builder to `amd64`, verifies the OpenAI tunnel-client checksum, runs without published inbound ports, drops all capabilities before adding only the startup minimum, and auto-restarts.
- The supplied OneDrive cloud artwork is packaged as the plugin logo while the existing square composer icon remains available for compact UI surfaces.

## Explicitly constrained coverage

- Business Graph Excel and organization-only sharing are mock-tested because the live account is personal.
- The isolated plus-address recipient was rejected by Graph; the harness verified that no grant existed, then completed live anonymous-link creation, audit, and exact revocation.
- Existing credentials, consent, and Keychain data remain untouched; forced login polling and Keychain deletion are excluded.
- Live recycle-bin restore remains excluded; native version restore passes live, while recycle-bin behavior passes mock and dry-run coverage.
