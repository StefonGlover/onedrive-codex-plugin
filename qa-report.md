# OneDrive 0.5.1 Release Gate Report

Decision: Pending — Entra registration, NAS OAuth rollout, ChatGPT Work validation, and correct hosted logo
Date: 2026-07-19
Generated: 2026-07-19T22:47:17Z
Tested source base commit: `87d94eb3df28c5c7a79409cbae238633ed992094`
Plugin version: `0.5.1+codex.20260719224717`
Tool contract: 84 exact tool names

## Current outcome

The `0.5.1` patch fixes both live Synology failures that prevented structured workbook reads: `EPERM: operation not permitted, chmod '/data'` and the follow-on `EACCES: permission denied, mkdir '/data/pycache'`. The Office helper no longer modifies the storage mount and keeps disposable Python bytecode in a private temporary directory outside `/data`. A focused real-XLSX fixture test, the full 157-check mock Graph/Office suite, the `nas9` non-writable `/data` mount test, immutable-cache parity across 52 packaged files, and the deployed `nas8` live workbook retest all pass. The exact `Personal/Documents/Career Development/QSE Job Tracker.xlsx` read completed through the Open XML backend with three worksheets, two tables, three charts, and 5,000 returned cells.

The pending OAuth build adds Entra protected-resource discovery, strict bearer-token validation, Graph on-behalf-of exchange, Streamable HTTP transport, per-tool `oauth2` scopes, runtime reauthorization challenges, and an HTTP-target Secure MCP Tunnel profile. The isolated OAuth integration test passes discovery, JWT signature/issuer/audience/time/scope checks, OBO exchange/cache behavior, all 84 OAuth descriptors, an unauthenticated challenge, and an authenticated call. This build is not yet a release Pass: it also needs the two Entra registrations, NAS OAuth deployment, a fresh ChatGPT Work host-loop result, and the correct hosted logo.

A live ChatGPT plugin-directory check at 2026-07-19T21:34:46Z confirmed that the current OneDrive listing still shows the generic plug. The existing app manager exposes name and description edits but no logo control. The New Plugin replacement flow does expose a PNG icon field plus OAuth/Mixed authentication; Chrome refused the automated file selection because the ChatGPT Chrome Extension cannot access file URLs. The correct square OneDrive cloud remains a mandatory replacement-app and visual-verification gate.

The No Auth descriptor repair passes both full 106-check source and immutable-cache betas with 96 passes, zero failures, ten explicit environment or safety blocks, exact coverage of all 84 tools, verified remote cleanup, and no isolated local residue. All 157 mock Graph checks pass, source/cache byte-mode-type-symlink parity passes, and the Synology `nas8` image is healthy on Node 20.20.2 with the persistent encrypted-auth and data mounts preserved.

A freshly registered No Auth app exposes `noauth` on all 84 actions and completes the read-only health check in regular Chat with five passes, one preset warning, and zero failures. The identical app is blocked by an incorrect expired-connection card in the Work surface before a tool call reaches the tunnel. Because the server is never invoked and the same app succeeds in Chat, this is recorded as a ChatGPT Work host limitation; selecting **Chat** is the verified workaround.

The first installed-build attempt hit the harness's default 10-second Graph timeout during managed Office backup restore. The harness hard-failed, deleted and verified absence of its exact isolated remote folder, and removed local work. A fresh complete run using the documented 60-second request ceiling then passed 96/10/0 with exact cleanup.

## Fixes validated

- Every MCP tool explicitly advertises `securitySchemes: [{ "type": "noauth" }]` and mirrors it under `_meta.securitySchemes`; ChatGPT's action inspector now shows the correct scheme on all 84 tools and the fresh app runs successfully in Chat.
- Orphaned `relativePath` arguments now fail closed instead of silently resolving scans or delta reads against the drive root.
- Doctor resolves every configured preset and warns with exact missing aliases instead of unconditionally reporting preset health.
- Recursive scans use bounded folder concurrency, report the effective concurrency, and exclude special non-file items from file-only results.
- Tool failures retain compatible text while adding stable machine-readable error metadata.
- The README now enumerates all 84 tools, including managed workspaces and change watches.
- The new doctor behavior is verified through the deployed NAS tunnel: authentication, profile, drive, and root access pass, while absent `Desktop` and `Pictures/Screenshots` targets are surfaced as explicit configuration warnings.
- The NAS rollout staged and verified all packaged files in a versioned app directory before switching the project to `nas8`; persistent encrypted auth and data mounts were preserved, and the exact QSE workbook now opens successfully through the live tunnel.

- Native version restore now uses Microsoft Graph's `restoreVersion` action and passes live plus mock verification.
- The beta harness uses one configurable Office Python runtime, cleans partial fixture setup, persists optional JSON reports, emits compact progress, and bounds child Graph request latency.
- Live search no longer multiplies transient retries across layers.
- Named-share testing rejects the owner's own mailbox and safely classifies a verified `sharingFailed` result only when no permission was created.
- Excel Open XML edits now preserve declared compatibility namespaces, reject undeclared `mc:Ignorable` prefixes, copy row formatting and height for appended table rows, and extend matching conditional-formatting and data-validation ranges.
- Excel `deleteTableRow` now compacts a bounded table row, preserves surrounding worksheet content, translates ordinary relative formulas, shrinks bounded single-column shared-formula groups without changing formulas or cached values, shifts native hyperlinks, shrinks matching ranges, preserves row formatting, and maintains Excel-required ascending cell order.
- Linux and NAS deployments use an AES-256-GCM encrypted token vault with a separately mounted owner-only key, atomic writes, symlink refusal, and strict permission validation.
- Synology startup now normalizes an uploaded encrypted token to `0600` before dropping privileges, fixing DSM/File Station's permissive upload mode without weakening vault checks.
- The DS923+ image pins the classic DSM builder to `amd64`, verifies the OpenAI tunnel-client checksum, runs without published inbound ports, drops all capabilities before adding only the startup minimum, and auto-restarts.
- The supplied square 256×256 OneDrive artwork is now used for the composer, primary, and dark-mode plugin icons; this removes the wide-banner aspect-ratio mismatch on plugin icon surfaces.

## Explicitly constrained coverage

- Business Graph Excel and organization-only sharing are mock-tested because the live account is personal.
- The isolated plus-address recipient was rejected by Graph; the harness verified that no grant existed, then completed live anonymous-link creation, audit, and exact revocation.
- Existing credentials, consent, and Keychain data remain untouched; forced login polling and Keychain deletion are excluded.
- Live recycle-bin restore remains excluded; native version restore passes live, while recycle-bin behavior passes mock and dry-run coverage.
- ChatGPT Work currently blocks this custom No Auth app with a false expired-connection card before tool dispatch. Regular Chat passes the same prompt and is the supported workaround pending a ChatGPT host correction.
- Synology administrative verification used HTTPS over the LAN fallback with insecure TLS certificate validation enabled after the Tailscale endpoint failed. The deployed service itself remains outbound-only, but DSM should be given a trusted certificate or certificate pinning to remove this operational warning.
