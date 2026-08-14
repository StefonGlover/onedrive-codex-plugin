# OneDrive 0.6.4 Parity Release Report

Decision: Pending — NAS60 nested-folder optimization and the refreshed ChatGPT 19-action OAuth app passed; the full live Microsoft Graph and installed-cache parity matrix remains pending
Date: 2026-08-14
Generated: 2026-08-14T20:35:27Z
Tested source base commit: `19029bff8ee6a525b695e3b7c808b91334696442`
Plugin version: `0.6.4+codex.20260813025943`
Focused source server version: `0.6.4+codex.20260813025943.chatgpt.fd16f4230688`
Live OAuth server version: `0.6.4+codex.20260813025943.chatgpt.eb2ef5cc3116`
Tool contract: 84 exact tool names; 19 focused ChatGPT tools
Packaged-content digest: `31a68740b900ee17a6cd56d0a260f2ee82b193e8f9ebc3ec1efc3f54542023f0`
NAS release-manifest SHA-256: `fecbd001a89c5a41e3c0053e0d5f48c9527ca7aaf0dbb22f7faffb984f7ef98e`
Server SHA-256: `1e82d65f417fca38daef6a95b7a1f85c8e403576756f7f7ff12a626d28e0a2ab`

## 2026-08-13 release candidate

The parity implementation is committed and the final frozen offline suite is green. It expands the hosted surface from 15 to 19 focused tools, adds bounded Office inspection/review, materialized downloads and previews, version/recent reads with native guarded restore, enterprise drive discovery/fetch, account-and-drive state isolation, and the associated availability and hosted-boundary hardening.

Current release-candidate evidence: mock Microsoft Graph 208/208, OAuth compatibility 169/169, full/focused/OAuth tool contracts 84/19/19, ChatGPT golden routing 19/19 with 15 ambiguity pairs, prepackage self-check 32/32, materialized resources 50/50, heavyweight admission 24/24, resources/read admission 17/17, bounded local state 3017/3017, and managed quota 11/11. NAS60 is deployed and the refreshed ChatGPT OAuth app passed both a general read-only canary and an exact nested-path workbook read; the full current-release live Graph and installed-cache evidence remain pending.

The detailed older sections below preserve prior production evidence as historical baseline and rollback context.

## 2026-08-14 NAS60 nested-folder deployment and ChatGPT refresh

NAS60 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas60-nested-folders-20260814` as `onedrive-chatgpt-nas:0.6.4-nas60-nested-folders-20260814`. Container `b65cdceb7421c6fb1438cd527ce11619645420a5b20cc00244c0e7b4fab0cf98` is running and healthy with zero restarts on image `sha256:4d7e7e8a066bcaa255e6ae96675205468c7f784697e53fe35d111acc4e07eafb`, Node 24.19.0, host networking, the preserved 2 GiB memory limit, and the existing persistent data/runtime mounts. The exact NAS59 source, compose identity, and image `sha256:0faf833f518e9361127bb95c07c9225934ea0665330f0c1f91a4b9b07f21bff7` remain available for rollback.

The source tree contained 92 packaged files; NAS staging added one release manifest, and all 93 remote files passed MD5-and-size verification. Public health, OAuth authorization-server metadata, protected-resource metadata, MCP `GET` 405 behavior, and MCP initialization passed. The live OAuth server version exactly matched the production-environment source contract at `0.6.4+codex.20260813025943.chatgpt.eb2ef5cc3116`, with 19 focused actions.

ChatGPT refreshed the existing mapped app `plugin_asdk_app_6a6995abb030819187d50d7080d4ae95` in place from `https://glovernas.tail5dbbc4.ts.net/mcp` and displayed `Actions refreshed.` All 19 expected actions remained present, and the refreshed `onedrive_open_files` action now documents known root-relative path resolution and a shared bounded filename fallback traversal. No duplicate app was created.

Two post-deployment reads passed without mutation. `onedrive_read_actions` returned one recent item through the existing OAuth connection. Then `onedrive_open_files` opened the known nested workbook path `Personal/Documents/Career Development/QSE Job Tracker.xlsx` in 4,654 ms. The focused mock regression separately proves that this path mode issues one direct metadata request with zero search and zero folder scans, while two filename index misses share each traversed folder once. The full live matrix remains pending, so the release decision remains Pending.

## 2026-08-13 NAS59 and ChatGPT canary

NAS59 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas59-parity-20260813-healthfix` as `onedrive-chatgpt-nas:0.6.4-nas59-parity-20260813`. Container `83b8cb26f3fd75fd3713dc71d3fe302dcde854638a6de8d83c3220c78eecad38` is running and healthy with zero restarts, Node 24.19.0, host networking, and loopback-only MCP and OAuth compatibility listeners on ports 3012 and 3011 respectively. Public health, OAuth authorization-server metadata, protected-resource metadata, and MCP initialization passed. The prior NAS58 image and source directory remain available for exact rollback.

The initial Container Manager cutover failed safely because this Synology kernel does not support the Compose `cpus`/NanoCPUs setting. The old healthy container remained available, the unsupported setting was removed, and the corrected immutable NAS59 image then deployed successfully with its 2 GiB memory and 256-PID limits intact.

ChatGPT refreshed the mapped OneDrive OAuth app from `https://glovernas.tail5dbbc4.ts.net/mcp`; the intended 19-action contract was present. A new-chat read-only `onedrive_office_capabilities` canary returned `addTableRow` with a schema-valid example: https://chatgpt.com/c/6a7d63fe-9dbc-83ea-bf20-90ee31309c17. No remote mutation was performed.

The source/noauth focused suffix `.37923d62a462` and live OAuth suffix `.d8d3679a1016` are deterministic variants of the same artifact: the OAuth security scheme is included in the advertised contract hash. Direct checks inside the deployed container confirmed exact source parity for `mcp/server.mjs`, `assets/chatgpt-icon.png`, and `.codex-plugin/plugin.json`.

This is deployment and connector-canary evidence only. The current-release full live Microsoft Graph matrix, complete focused live exercise, installed-cache parity, doctor/tenant checks, and cleanup evidence remain pending, so the release decision remains Pending.

## Outcome

The OneDrive plugin passed the complete offline suite, the new read-only ranked-search beta against the connected personal OneDrive through the exact 15-tool ChatGPT Work profile, a prior fresh 84-tool live Microsoft Graph run, direct production calls for every focused ChatGPT function, and current production UI workflows in both ChatGPT Work and standard Chat. The new search beta made no remote mutations and used isolated local state.

The source ranked-search beta `codex-beta-work-20260809T062753Z` and installed-cache rerun `codex-beta-work-installed-20260809T063041Z` each passed 5/5 checks. In both runs, all three representative queries returned the intended target at rank 1, for exact-at-one 3/3, MRR@10 1.0, and zero known unrelated results in the top five. Both runs also fetched readable text from `HACCP Study Guide.pdf`. Evidence: `work/qa-artifacts/codex-beta-work-20260809T062753Z.json` and `work/qa-artifacts/codex-beta-work-installed-20260809T063041Z.json`.

The beta found and resolved three release issues:

- The shared Tailscale HTTPS router had lost the root handler needed for OneDrive `/authorize` and `/token`, while `/mcp` remained present. This caused reconnect 404s and ChatGPT internal plugin errors. The root handler was restored to `127.0.0.1:3011`, `/mcp` remains routed to `127.0.0.1:3012/mcp`, all other app prefixes were preserved, OAuth was reconnected, and ChatGPT reads and writes passed afterward.
- The live harness called rename, move, and copy commits without their now-required preview proofs. It now previews each action and refuses to continue without concrete verified commit evidence. The repaired live run passed all three steps.
- CI Python syntax checks generated `__pycache__` inside the source tree and then failed the package-residue guard. CI now writes bytecode under the runner temporary directory.

The ranked source was deployed to the NAS build context `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas57-chatgpt-ranked-20260809`, built successfully, and started with exit code 0. Synology Container Manager retained the existing project image tag `onedrive-chatgpt-nas:0.6.4-nas56-chatgpt-etag` when it reloaded the compose, but rebuilt that tag from the ranked nas57 source context. The prior source was preserved at `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas56-chatgpt-etag-20260808-rollback-20260809` for rollback. Public health, OAuth authorization-server metadata, protected-resource metadata, and the MCP route all passed after restart.

## Live coverage

### Production ranked deployment and connector refresh

ChatGPT refreshed the OneDrive connector from `https://glovernas.tail5dbbc4.ts.net/mcp` and reported `Success: Actions refreshed.` OAuth remained the supported and active authorization method. A post-restart public health check returned healthy OAuth and MCP status.

The production read-only ranking beta passed on both current ChatGPT surfaces:

- Work: `Qaldris` returned the Qaldris folder first, followed by two related launch-kit manifests; `HACCP Study Guide.pdf` and `Digital Quality Management Insights (1).docx` each returned the exact file as the only result. Evidence: https://chatgpt.com/c/6a78270d-4cd8-83ea-83e5-46ac7fd0e750
- Standard Chat: the connected OneDrive plugin invoked `onedrive_read_actions`; the same three queries produced the same ordering and exact matches. A second exact HACCP search after the final NAS restart also passed in 3,557 ms. Evidence: https://chatgpt.com/c/6a78276b-1634-83ea-ac49-d5e568709f98

These UI betas were read-only and made no OneDrive mutations.

### Focused 15-tool production connector

All focused functions were exercised against the connected personal OneDrive:

- Batched list, search, item-info, and permission reads
- Bounded content fetch and exact-name open
- Folder creation
- Guarded text creation and patching
- Batched rename, move, copy, anonymous-link creation, and permission revocation
- Named-recipient invite preview plus live Microsoft Graph error-state audit
- Local-file upload
- Office capability discovery
- Guarded Word batch transformation with backup and remote package verification
- Word-to-PDF export
- Recycle deletion and exact deleted-item restore
- Final recoverable cleanup and active-path absence verification

The isolated root was `ChatGPT OneDrive Plugin Beta Test chatgpt-direct-20260808-2236`. It was recycled at the end and no active fixture root or anonymous link remained.

### Full 84-tool live harness

Fresh run `codex-beta-20260808t2243z-fix` completed 106 checks:

- Pass: 97
- Fail: 0
- Blocked by account/environment: 9
- Exact non-blocked tool contract exercised: pass
- Final root cleanup: pass

Evidence: `work/qa-artifacts/live-codex-beta-20260808t2243z-fix.json`

### ChatGPT Work UI

The final Work workflow passed 15/15 checkpoints: OAuth-backed root read, root and nested-folder creation, guarded text write and exact readback, rename, move, copy, anonymous-link creation, permission inspection, exact link revocation, recycle deletion, deleted-item restore, restored-content verification, final root recycling, and expected `itemNotFound` at the former active path.

Evidence: https://chatgpt.com/c/6a77ad45-6fec-83ea-acae-a930238c07f2

### ChatGPT standard Chat UI

Ordinary Chat selected the connected OneDrive plugin and passed root listing, isolated folder creation, guarded text write, exact readback including the trailing newline, the in-product permission prompt, recycle cleanup, and expected `itemNotFound` at the former active path.

Evidence: https://chatgpt.com/c/6a77b265-cac4-83ea-a802-65c093afe437

## Offline verification

- Mock Microsoft Graph: 188/188
- Ranked ChatGPT Work search: source and installed-cache live read-only betas each passed 5/5 checks; exact-at-one 3/3; MRR@10 1.0; unrelated top-five 0
- Full MCP contract: 84 tools, 339,137 descriptor bytes
- Focused ChatGPT contract: 15 tools, 26,526 bytes; 92.2% reduction
- Focused OAuth contract: 15 tools, 28,656 bytes
- Golden routing: 15/15 prompts and 8 ambiguity pairs
- OAuth HTTP integration: pass in full and focused profiles
- OAuth compatibility: 169/169
- Semantic anchors: 6/6
- Text patch safety: 6/6
- Word operation coverage: 21/21
- Excel operation coverage: 33/33
- PowerPoint operation coverage: 25/25
- Genuine Office packages reopened and rendered: DOCX, XLSX, and PPTX pass
- Storage-root and private-permission tests: pass
- Plugin package guard, immutable versioned-cache installation, and installed-build parity: pass across 62 package entries

## External constraints, not plugin defects

- Named-recipient invitation: a user-controlled Gmail plus alias was tested in both silent and email modes. Microsoft Graph returned `sharingFailed`; a permission audit proved no grant or partial state was created. A distinct non-owner Microsoft recipient is required to demonstrate the success path live. The success and cleanup paths remain covered by the mock suite.
- Business Graph Excel sessions: unavailable on the connected personal OneDrive; fully mock-tested.
- Organization-scoped sharing: unavailable on the connected personal OneDrive; fully mock-tested.
- Forced device-code polling and credential deletion were intentionally excluded to preserve the healthy production credential; their safe paths are mock-tested.

The detailed capability matrix is in `work/qa-artifacts/capability-matrix-20260807.md`.
