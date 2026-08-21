# OneDrive 0.6.5 Integrity and Production Release Report

Decision: Pending — 0.6.5 offline gates pass; immutable NAS69 deployment, same-app refresh, signed-in ChatGPT beta, cache parity, and exact fixture cleanup remain
Date: 2026-08-21
Generated: 2026-08-20T21:25:15Z
Tested source base commit: `8eaf5e4ae91c737d913ff9b8169f1a8a98c8fd0c`
Plugin version: `0.6.5+codex.20260820212515`
Focused source server version: `0.6.5+codex.20260820212515.chatgpt.bf50db0b0ca8`
Current live NAS OAuth server version: `0.6.4+codex.20260819032617.chatgpt.a8e1a22814f9`
Tool contract: 84 exact tool names; 21 focused ChatGPT tools
Packaged-content digest: `4d8b4fffe5437ca763bd01198f0226881814ef53a93a0b0b62aa01808e3e95ab`
NAS release-manifest SHA-256: pending immutable NAS68 staging
Server SHA-256: `106dc6e87b18ca97dd2c304a2c90acfc5d15214f1bca6292227028a7cf36e26f`

## 2026-08-21 NAS69 Excel integrity and production release candidate

Release `0.6.5+codex.20260820212515` adds automatic Excel formula/reference integrity reporting and a fail-closed post-edit gate. It detects formula error tokens, stored error cells, missing sheet/table references, broken defined names, static circular references, external links, volatile formulas, calculation mode, and cache coverage. Sheet renames now rewrite dependent formulas, defined names, tables, and charts. Personal/OpenXML work stays honest with `calculationVerified: false`; supported Business/SharePoint workbooks can invoke Microsoft's Graph calculation engine in a persistent workbook session.

Template requests now resolve, inspect, preview-copy, commit-copy, edit, and re-inspect the exact reference file instead of creating a blank package. Standard Chat metadata routing explicitly enriches exact size, MIME, and modified time through item-info. Production health exposes bounded release, error/throttle, and p50/p95 tool latency counters without tenant content. The release adds an external OAuth/MCP production canary, a repeated-search latency benchmark, exact dependency/base-image checks, deterministic CycloneDX source inventory, and commit-pinned Trivy image/SBOM CI.

Offline evidence is green at 213/213 mock Graph checks, including real Business workbook calculation routing; 169/169 OAuth compatibility checks; 21/21 golden prompts plus three cross-tool workflows and 18 ambiguity pairs; all 80 Office operations; the Office security corpus; native-library reopen; and LibreOffice render without repair diagnostics. The source-verified production image also has zero fixed High/Critical findings across the Debian runtime, Node package metadata, pinned Python packages, and rebuilt Go tunnel binary in the local Trivy gate. NAS69 downloads the exact upstream tunnel commit archive from GitHub codeload, verifies an independently pinned SHA-256, applies the reviewed dependency patches, and performs a transparent CGO-free, trimmed, VCS-disabled Go build with the version and source commit embedded. It additionally refuses any inherited Compose override that attempts to downgrade the reviewed tunnel client below 0.0.12. The focused contract remains 21 tools at 35,853 bytes without OAuth and 38,835 bytes with OAuth, under its 38 KiB cap. NAS69 deployment and signed-in ChatGPT beta are still pending, so this report does not yet claim release completion.

## 2026-08-18 NAS65 search and native Office creation release

This release exposes the MCP-standard `search` action in the focused ChatGPT profile, adds authentication-scoped 60-second search snapshots with mutation invalidation, enriches compact result metadata, and adds guarded native creation of genuine Word, Excel, and PowerPoint packages. The new creation flow generates and strictly reopens each Open XML package before upload, binds confirmation to the destination and specification, refuses to mint a replacement token without matching remote identity and eTag evidence, and uses create-only semantics if a previously absent destination becomes occupied before commit.

The focused contract is 21 actions while the full contract remains exactly 84 tools. Offline evidence is green at 213/213 mock Graph checks, 169/169 OAuth compatibility checks, 21/21 ChatGPT routing prompts with 18 ambiguity pairs, five remote workflow skills with 16 resources, and all 80 typed Office edit operations. The versioned cache is installed at `$CODEX_HOME/plugins/cache/personal/onedrive/0.6.4+codex.20260819032617`; all 92 packaged files match source across bytes, modes, types, and symlink targets without overwriting an older cache.

NAS65 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas65-search-office-create-20260818` as `onedrive-chatgpt-nas:0.6.4-nas65-search-office-create-20260818`, image `sha256:c469c000e69ebbf09f0808f211254f8d7edc30e07fa3869aa7fd7cd56ae25431`. Container `41a80b4a96895f321de510ab62634fca31d4c3cd22f7a00eb50e8e097471ba70` is running and healthy with zero restarts and zero health-check failures. Public health, OAuth metadata, protected-resource metadata, expected MCP `GET` 405 behavior, unauthenticated tool rejection, and MCP 2025-11-25 initialization all pass. The immutable NAS64 source and image remain available for exact rollback.

The existing ChatGPT app `plugin_asdk_app_6a6995abb030819187d50d7080d4ae95` was refreshed in place with the same OAuth connection and now exposes all 21 focused actions; no duplicate app was created. A signed-in [ChatGPT Work beta](https://chatgpt.com/c/6a852977-fa10-83ea-82e9-eb890b5aeb0e) passed guarded native Word preview and creation, remote package validation, two identical standard searches with a second-hit `scoped_memory` cache result at age 4,357 ms and TTL 60,000 ms, fetch readback by stable item ID, and an occupied-destination refusal that issued no preview token. A separate [standard Chat beta](https://chatgpt.com/c/6a852ad3-c314-83ea-8de5-181eaece2b02) selected the connected OneDrive app and returned the exact filename, stable item ID, modified timestamp, and MIME type through MCP-standard search without mutation; its generic answer renderer did not surface the size or cache fields that Work exposed.

The sole disposable fixture, `Codex NAS65 Native Office Search Cache Test 20260818.docx`, was identity-checked and moved to the OneDrive recycle bin through a fresh guarded preview. Post-cleanup direct lookup fails, standard search no longer returns the stable item ID, and the deletion invalidated the scoped search cache. The cleanup is recoverable and was not a permanent deletion.

Overall, this connector now exceeds the Google Drive plugin for Microsoft-first document work through direct validated Office creation, 80 typed Office mutations, Office inspection/review/render/export, version recovery, guarded write proofs, MCP-standard search/fetch, and private self-hosted OAuth deployment. It does not claim a provider-native Google Docs comment-thread equivalent: Microsoft Graph `driveItem` exposes no comparable OneDrive comment service, so Word-package comments and Office review remain the honest Microsoft-native substitute.

## 2026-08-16 NAS64 workflow, latency, and materialized-resource release

Release `0.6.4+codex.20260816170509` advertises the five bundled OneDrive workflow skills through the MCP skills extension for import during **Scan Tools**, removes anonymous object-union rendering from four focused Office selectors, returns explicit not-applicable evidence for enterprise-only reads on personal accounts, and adds a subject-scoped 60-second list snapshot that is invalidated after successful mutations. A NAS63 live canary exposed a missing unprivileged `/data/materialized-resources` startup directory; NAS64 fixes that ownership contract and adds a regression assertion. NAS64 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas64-materialized-resources-20260816` as `onedrive-chatgpt-nas:0.6.4-nas64-materialized-resources-20260816`, image `sha256:982828dcd793c27238682eceddbaf1d01cb68d721c41e94e391ed07981341bb5`. Container `c1d1fcb89fc598609068eb9936bff68caaaff930a2597e3d55c6f994096d20d7` is running and healthy with zero restarts. NAS63 remains the exact rollback source and image.

Offline evidence is green: mock Microsoft Graph 212/212, OAuth compatibility 169/169, full/focused/OAuth tool contracts 84/19/19, ChatGPT golden routing 19/19, five remote skills with 16 digest-verified resources, all 80 advertised Office operations, hosted-boundary checks, and the remaining utility/security suites. The safe live focused beta passed 5/5 without remote mutation; all three representative queries ranked the intended result first with MRR@10 of 1.0. Public health, OAuth metadata, protected-resource metadata, MCP route behavior, initialize, skill scanning, materialized download, and preview rendering all pass on NAS64. The versioned cache is installed at `$CODEX_HOME/plugins/cache/personal/onedrive/0.6.4+codex.20260816170509`; all 92 package files match source across bytes, modes, types, and symlink targets.

The existing ChatGPT app `plugin_asdk_app_6a6995abb030819187d50d7080d4ae95` was refreshed in place with the same OAuth connection and 19 actions; no duplicate app was created. The refreshed `onedrive_office_inspect` and `onedrive_office_review` schemas expose stable named properties without top-level `anyOf`. A signed-in Work conversation passed four read-only prompts: cold root listing (3,139 ms), the same list from `scoped_memory` (0 ms, age 17,327 ms, TTL 60,000 ms), enterprise discovery returning `not_applicable`/`personal_account` with no personal-drive fallback, and `Dashboard Engine!A1:B5` inspection returning the stable workbook name/eTag, hidden sheet state, 10 cells, and three formulas. Evidence: https://chatgpt.com/c/6a81f0f9-efb8-83ea-9eec-4d8833873981. No OneDrive data was changed.

The current-release live scope intentionally excluded a fresh external recipient invitation and a new recycle/delete matrix because those operations affect other people or cloud data. Their unchanged contracts retain the immediate-predecessor full live baseline and are covered by the current 212-check mock suite. Current NAS64 live coverage includes safe Graph reads, Office inspection/review, materialized download, preview rendering, public OAuth/MCP discovery, remote-skill digest verification, and signed-in ChatGPT orchestration.

## 2026-08-14 NAS62 filename-only hyperlink deployment

Release `0.6.4+codex.20260814225653` makes the model-visible result from `onedrive_open_files` a Markdown hyperlink whose visible text is only the resolved filename. The canonical provider URL, stable item ID, extracted text, and metadata remain in structured content. NAS62 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas62-filename-links-20260814` as `onedrive-chatgpt-nas:0.6.4-nas62-filename-links-20260814`. Container `29b06e432925b5658a0dc0df7cdcfef125a7bce124735e111a706d7690a71967` is running and healthy with zero restarts on immutable image `sha256:e00ea47b54fa28535d6d537e32e4e95272ee33d1222851a5f479c2728e63598f`. NAS61 remains the exact rollback source and image.

All 92 package files plus the release manifest were staged under the immutable NAS62 source directory and individually verified by size and MD5. Public health, OAuth metadata, protected-resource metadata, MCP route behavior, and initialization passed; the live OAuth server advertises `0.6.4+codex.20260814225653.chatgpt.b8b8204c94fd`. The existing mapped app `plugin_asdk_app_6a6995abb030819187d50d7080d4ae95` was refreshed in place from the existing production URL and displayed `Actions refreshed.` with the 19-action focused contract. No duplicate app was created.

The fresh ChatGPT Work canary searched for `QSE Job Tracker.xlsx` and returned a single decorated hyperlink whose visible text was exactly `QSE Job Tracker.xlsx`; no raw provider URL appeared in the assistant response. Clicking the filename opened ChatGPT's external-link confirmation for the OneDrive provider URL. Evidence: https://chatgpt.com/c/6a7fa0df-bda8-83ea-9e70-6ffe491bd9aa. The probe was read-only and made no remote mutation.

Offline evidence is green: mock Microsoft Graph 210/210, full/focused/OAuth tool contracts 84/19/19, ChatGPT golden routing 19/19, hosted-boundary checks, and prepackage self-check 32/32. The versioned personal cache was installed without overwriting older caches, and all 92 packaged files passed byte, mode, type, and symlink-target parity. The immediately preceding NAS61 release completed both fresh 84-tool source and installed-cache live matrices; NAS62 changes only model-visible link presentation, covered by focused regressions and the live ChatGPT canary above.

## 2026-08-13 release candidate

The parity implementation is committed and the final frozen offline suite is green. It expands the hosted surface from 15 to 19 focused tools, adds bounded Office inspection/review, materialized downloads and previews, version/recent reads with native guarded restore, enterprise drive discovery/fetch, account-and-drive state isolation, and the associated availability and hosted-boundary hardening.

Final release evidence: mock Microsoft Graph 210/210, full/focused/OAuth tool contracts 84/19/19, ChatGPT golden routing 19/19 with 15 ambiguity pairs, hosted-boundary checks, and prepackage self-check 32/32. The new read-only `urls` mode validates OneDrive/SharePoint hosts, resolves each link through Microsoft Graph share addressing, reads from the exact case-preserved source drive, rejects mixed or untrusted selectors before Graph, and does not echo access-bearing input links. NAS61, the existing ChatGPT app, and both fresh full live matrices are now verified.

The detailed older sections below preserve prior production evidence as historical baseline and rollback context.

## 2026-08-14 NAS61 direct-link deployment and ChatGPT beta

NAS61 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas61-link-open-20260814` as `onedrive-chatgpt-nas:0.6.4-nas61-link-open-20260814`. Container `1e0862a470b3fe4e391dd3bfbc601101f339e95c15b57f111ae0091524e1c90e` is running and healthy with zero restarts on immutable image `sha256:06c4392e2f41fa87a26eef592dbb425ed47e02bc8e0eb9160734a038c0e2879d`. Public health, OAuth authorization-server metadata, protected-resource metadata, and expected MCP `GET` 405 behavior passed. Exact NAS60 source and image `sha256:4d7e7e8a066bcaa255e6ae96675205468c7f784697e53fe35d111acc4e07eafb` remain the rollback target.

All 92 packaged source files plus the release manifest were staged under the new immutable NAS61 directory and individually verified by MD5 and size. The release-manifest SHA-256 is `89fc9cdabb96601455c9db68e0231f86bafdcb4756e40dd752cf7092eae339ff`; the deployed server SHA-256 is `3960272b1f55e5c95331e0cdaaaee43b8e2cec8d659b87bb3c66aa8bab107dcd`.

ChatGPT refreshed the existing mapped app `plugin_asdk_app_6a6995abb030819187d50d7080d4ae95` in place from `https://glovernas.tail5dbbc4.ts.net/mcp` and displayed `Actions refreshed.` The refreshed `onedrive_open_files` action exposes bounded `names` and `urls` arrays while the focused surface remains 19 actions. No duplicate app was created.

The new ChatGPT Work conversation passed three read-only probes: the exact nested workbook path resolved to its stable item ID and provider-observed OneDrive URL; that observed URL then opened directly and was classified as URL input with the same stable identity; and `https://example.com/not-a-onedrive-link` was rejected as untrusted. Evidence: https://chatgpt.com/c/6a7f939f-1ee4-83ea-96a1-d83172d0a3d0.

The current-source and independently installed-cache full live matrices each completed 106 checks with 97 passed, zero failed, nine explicitly blocked, and 83 of 84 tools exercised. Runtime was 248,496 ms for `codex-beta-nas61-source-20260814` and 284,234 ms for `codex-beta-nas61-installed-20260814`. Both isolated remote fixture roots and local work directories were removed. A final stale-root audit found zero cleanup candidates, and installed parity matched all 92 packaged files across bytes, modes, file types, and symlink targets without overwriting an older cache.

## 2026-08-14 NAS60 nested-folder deployment and ChatGPT refresh

NAS60 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas60-nested-folders-20260814` as `onedrive-chatgpt-nas:0.6.4-nas60-nested-folders-20260814`. Container `b65cdceb7421c6fb1438cd527ce11619645420a5b20cc00244c0e7b4fab0cf98` is running and healthy with zero restarts on image `sha256:4d7e7e8a066bcaa255e6ae96675205468c7f784697e53fe35d111acc4e07eafb`, Node 24.19.0, host networking, the preserved 2 GiB memory limit, and the existing persistent data/runtime mounts. The exact NAS59 source, compose identity, and image `sha256:0faf833f518e9361127bb95c07c9225934ea0665330f0c1f91a4b9b07f21bff7` remain available for rollback.

The source tree contained 92 packaged files; NAS staging added one release manifest, and all 93 remote files passed MD5-and-size verification. Public health, OAuth authorization-server metadata, protected-resource metadata, MCP `GET` 405 behavior, and MCP initialization passed. The live OAuth server version exactly matched the production-environment source contract at `0.6.4+codex.20260813025943.chatgpt.eb2ef5cc3116`, with 19 focused actions.

ChatGPT refreshed the existing mapped app `plugin_asdk_app_6a6995abb030819187d50d7080d4ae95` in place from `https://glovernas.tail5dbbc4.ts.net/mcp` and displayed `Actions refreshed.` All 19 expected actions remained present, and the refreshed `onedrive_open_files` action now documents known root-relative path resolution and a shared bounded filename fallback traversal. No duplicate app was created.

Two post-deployment reads passed without mutation. `onedrive_read_actions` returned one recent item through the existing OAuth connection. Then `onedrive_open_files` opened the known nested workbook path `Personal/Documents/Career Development/QSE Job Tracker.xlsx` in 4,654 ms. The focused mock regression separately proves that this path mode issues one direct metadata request with zero search and zero folder scans, while two filename index misses share each traversed folder once.

The current-source full live matrix then exercised 83 of the exact 84 tools against the isolated folder `Codex OneDrive Plugin Beta Test codex-beta-nas60-source2-20260814`. It completed 106 checks in 242,771 ms: 97 passed, zero failed, and nine were explicitly blocked by personal-account, provider, or credential-preservation limitations. The only uncalled contract tool was destructive device-code polling. The harness revoked its temporary sharing grants, deleted the exact test folder, verified it absent, and removed its local isolated work directory.

The independently installed versioned cache repeated that result against `Codex OneDrive Plugin Beta Test codex-beta-nas60-installed-20260814`: 106 checks in 264,895 ms, with 97 passed, zero failed, nine explicitly blocked, and 83 of 84 tools exercised. Its remote root and local isolated work were removed. A final stale-root discovery returned zero beta folders, and the source/cache verifier matched all 92 packaged files across bytes, modes, file types, and symlink targets without overwriting an older cache.

## 2026-08-13 NAS59 and ChatGPT canary

NAS59 is deployed from `/volume1/docker/onedrive-chatgpt/app-0.6.4-nas59-parity-20260813-healthfix` as `onedrive-chatgpt-nas:0.6.4-nas59-parity-20260813`. Container `83b8cb26f3fd75fd3713dc71d3fe302dcde854638a6de8d83c3220c78eecad38` is running and healthy with zero restarts, Node 24.19.0, host networking, and loopback-only MCP and OAuth compatibility listeners on ports 3012 and 3011 respectively. Public health, OAuth authorization-server metadata, protected-resource metadata, and MCP initialization passed. The prior NAS58 image and source directory remain available for exact rollback.

The initial Container Manager cutover failed safely because this Synology kernel does not support the Compose `cpus`/NanoCPUs setting. The old healthy container remained available, the unsupported setting was removed, and the corrected immutable NAS59 image then deployed successfully with its 2 GiB memory and 256-PID limits intact.

ChatGPT refreshed the mapped OneDrive OAuth app from `https://glovernas.tail5dbbc4.ts.net/mcp`; the intended 19-action contract was present. A new-chat read-only `onedrive_office_capabilities` canary returned `addTableRow` with a schema-valid example: https://chatgpt.com/c/6a7d63fe-9dbc-83ea-bf20-90ee31309c17. No remote mutation was performed.

The source/noauth focused suffix `.37923d62a462` and live OAuth suffix `.d8d3679a1016` are deterministic variants of the same artifact: the OAuth security scheme is included in the advertised contract hash. Direct checks inside the deployed container confirmed exact source parity for `mcp/server.mjs`, `assets/chatgpt-icon.png`, and `.codex-plugin/plugin.json`.

This historical NAS59 section is deployment and connector-canary evidence only; the current NAS61 deployment, source and installed-cache matrices, doctor/tenant checks, and cleanup audit provide the release-complete evidence above.

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

## Offline and release verification

- Mock Microsoft Graph: 210/210
- Current source live Microsoft Graph: 97 passed, 0 failed, 9 explicitly blocked; 83/84 tools exercised; exact cleanup passed
- Installed-cache live Microsoft Graph: 97 passed, 0 failed, 9 explicitly blocked; 83/84 tools exercised; exact cleanup passed
- Ranked ChatGPT Work search: source and installed-cache live read-only betas each passed 5/5 checks; exact-at-one 3/3; MRR@10 1.0; unrelated top-five 0
- Full MCP contract: 84 exact tools
- Focused ChatGPT contract: 19 tools
- Focused OAuth contract: 19 tools
- Golden routing: 19/19 prompts and 15 ambiguity pairs
- OAuth HTTP integration: pass in full and focused profiles
- OAuth compatibility: 169/169
- Semantic anchors: 6/6
- Text patch safety: 6/6
- Word operation coverage: 22/22
- Excel operation coverage: 33/33
- PowerPoint operation coverage: 25/25
- Genuine Office packages reopened and rendered: DOCX, XLSX, and PPTX pass
- Storage-root and private-permission tests: pass
- Plugin package guard, immutable versioned-cache installation, and installed-build parity: pass across 92 packaged files
- Final cleanup audit: zero active beta roots, zero temporary permissions or anonymous links, and zero isolated local residue

## External constraints, not plugin defects

- Named-recipient invitation: a user-controlled Gmail plus alias was tested in both silent and email modes. Microsoft Graph returned `sharingFailed`; a permission audit proved no grant or partial state was created. A distinct non-owner Microsoft recipient is required to demonstrate the success path live. The success and cleanup paths remain covered by the mock suite.
- Business Graph Excel sessions: unavailable on the connected personal OneDrive; fully mock-tested.
- Organization-scoped sharing: unavailable on the connected personal OneDrive; fully mock-tested.
- Forced device-code polling and credential deletion were intentionally excluded to preserve the healthy production credential; their safe paths are mock-tested.

The detailed capability matrix is in `work/qa-artifacts/capability-matrix-20260807.md`.
