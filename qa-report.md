# OneDrive 0.5.1 Release Gate Report

Decision: Pending — Entra registration, NAS OAuth rollout, and ChatGPT Work validation
Date: 2026-07-20
Generated: 2026-07-20T15:16:15Z
Tested source base commit: `8f5ae49c61e0beb79226cd2a038b88b456235876`
Plugin version: `0.5.1+codex.20260719224717`
Tool contract: 84 exact tool names

## Current outcome

The final ChatGPT/NAS pass is deployed as `onedrive-chatgpt-nas:0.5.1-nas13`. The focused contract remains 19 tools and is now 33,669 bytes (90% smaller than the full 84-tool/335,838-byte contract), with 484 bytes of routing instructions. All 165 mocked Graph checks and all 19 golden prompt cases pass. Rename, move, and copy now require an account-scoped, operation-bound, single-use preview token in addition to confirmation and exact identity for live calls. DSM built image `ea810e8b9f6c`, started it with exit code 0, and reports all services healthy.

The canonical ChatGPT app remains **OneDrive** (`asdk_app_6a5e2416985481918d0f6c68785da2c4`); metadata refresh succeeded without recreating the app, and the stored local OneDrive image is still displayed. In the same ChatGPT conversation, the post-deploy beta returned `$12,325 — excel`, read one permission with no sharing-link permission, and an exact-ID `onedrive_rename` dry-run returned `dryRun`, `confirmed`, `wouldRename`, `newName`, `previewToken`, and `previewTokenExpiresAt` with no error. No file or permission was changed in this final pass. Host orchestration still took about 39 seconds for the multi-step read/permissions/preview prompt even though the plugin contract and warm read paths are compact.

One ChatGPT host-safety defect remains outside the plugin: during the earlier isolated CRUD fixture, clicking **Deny** on the host permission dialog did not prevent the folder/file mutation, and the consent UI displayed unrelated PII categories. The plugin's own scoped preview, identity, confirmation, audit, and cleanup protections remain enabled; this host behavior should be reported to OpenAI. The isolated fixture was subsequently moved to the recycle bin and its exact folder ID was verified.

The ChatGPT workbook-content failure is fixed in the tested source. Standard `fetch` now downloads modern Office packages once and renders bounded Word paragraphs/tables, Excel worksheet names/cells/formulas, and PowerPoint slides/shapes instead of relying on Microsoft Graph `format=text`. The exact `2026 Family Budgeting.xlsx` failure was reproduced in Chrome, the refreshed developer-mode app then returned a real worksheet-by-worksheet summary with exact figures, and the plugin logged `previewSource: office-openxml` with 40,679 bytes returned in 3.534 seconds on the cold read. A final warm run used only one `search` (1.052 seconds, zero Graph search calls) and one `fetch` (1.6 milliseconds from the content index); the ChatGPT host completed the detailed answer in 79.742 seconds.

The live ChatGPT contract now has 19 focused tools and 32,446 bytes, 90.3% smaller than the unchanged 84-tool/335,403-byte full contract. Redundant high-volume Office read tools were removed from the ChatGPT surface after the live host called `onedrive_excel_get_workbook` twice, produced 12.9 MB responses, and remained unfinished after 110 seconds. Standard `fetch` is now the single ChatGPT read path. The final live runs made no redundant Office call.

The current source hardens that focused path without changing the 19-tool surface. Its 32,446-byte descriptor set gives every tool a discriminative `Use this when` description and bounded ChatGPT invocation status, corrects the fetch-first Office-edit instructions, returns representative 32 KiB previews with sequential 64 KiB `fetch` continuation IDs, serves high-confidence matches from caches up to 24 hours old while revalidating in the background, and validates stale indexed content by ETag/cTag before reuse. The 19-case metadata golden-prompt gate, six ambiguity pairs, and all 164 mocked Graph checks pass. The exact source was deployed as `onedrive-chatgpt-nas:0.5.1-nas11`, came up healthy with server hash `5972c6de2076`, and was refreshed into the canonical 19-tool ChatGPT app.

The refreshed metadata fixes the live PDF search-to-fetch handoff by explicitly telling ChatGPT that every search result contains an opaque ID and to pass it unchanged to `fetch`. A clean regular-Chat run successfully returned `#2-200-OSLLenderChangeAmendment (2022_03_25 23_08_06 UTC).pdf`, preview source `local-pdf`, and heading `Amendment to Agreement`. The cold plugin calls took 1.323 seconds for stale-local search and 2.609 seconds for PDF extraction with zero Graph search calls. A warm repeat took 802 milliseconds for fresh-local search and 1.604 seconds for cached-metadata fetch. ChatGPT itself took roughly 35 seconds on the cold run, confirming that most remaining wall time is host orchestration rather than plugin execution.

Common file handling now includes direct text/code/CSV/TSV/JSON/XML/Markdown reads plus bounded local extraction for PDF, RTF, OpenDocument, EPUB, legacy `.doc`/`.xls`/`.ppt`, and common images. The Synology image installs Poppler, Tesseract, and catdoc. Dependency-free RTF, OpenDocument, and EPUB tests pass; the full mock Graph suite passes 164/164, including cold and warm structured Excel fetches, stale-index validation, progressive continuations, background search revalidation, and integrated RTF extraction.

The `0.5.1` patch fixes both live Synology failures that prevented structured workbook reads: `EPERM: operation not permitted, chmod '/data'` and the follow-on `EACCES: permission denied, mkdir '/data/pycache'`. The Office helper no longer modifies the storage mount and keeps disposable Python bytecode in a private temporary directory outside `/data`. A focused real-XLSX fixture test, the full 157-check mock Graph/Office suite, the `nas9` non-writable `/data` mount test, immutable-cache parity across 52 packaged files, and the deployed `nas8` live workbook retest all pass. The exact `Personal/Documents/Career Development/QSE Job Tracker.xlsx` read completed through the Open XML backend with three worksheets, two tables, three charts, and 5,000 returned cells.

The pending OAuth build adds Entra protected-resource discovery, strict bearer-token validation, Graph on-behalf-of exchange, Streamable HTTP transport, per-tool `oauth2` scopes, runtime reauthorization challenges, and an HTTP-target Secure MCP Tunnel profile. The isolated OAuth integration test passes discovery, JWT signature/issuer/audience/time/scope checks, OBO exchange/cache behavior, all 84 OAuth descriptors, an unauthenticated challenge, and an authenticated call. This build is not yet a release Pass: it still needs the two Entra registrations, NAS OAuth deployment, and a fresh ChatGPT Work host-loop result.

The canonical developer-mode app, **OneDrive** (`asdk_app_6a5e2416985481918d0f6c68785da2c4`), was created with the exact local `assets/chatgpt-icon.png` before any metadata refresh and later renamed from its temporary `OneDrive Fast` label after the stale name conflicts were removed. The uploaded asset is a 256×256 PNG, 2,276 bytes, SHA-256 `b9db1f911c59c34ce12cdfdfbae1a6b9933b140e3b024068a8df6ddac43fe5e1`. The app still displays the same image after the nas11 metadata refresh and rename. Both obsolete OneDrive developer registrations were permanently deleted, and their former plugin URLs now fail to load, leaving one installed OneDrive entry in the catalog.

The No Auth descriptor repair previously passed both full 106-check source and immutable-cache betas with 96 passes, zero failures, ten explicit environment or safety blocks, exact coverage of all 84 tools, verified remote cleanup, and no isolated local residue. The current source passes all 164 mock Graph checks. The Synology and ChatGPT rollouts now match this source; immutable Codex-cache parity remains a separate pending distribution step.

The NAS project retained the previous rollback directories and now also stores the exact 66-file nas11 source archive `app-0.5.1-nas11.zip`, SHA-256 `4343d3cb15342389cca9b72ccd2210e28ded0504a4f097751dcbdbd8cd18b756`.

The canonical No Auth app exposes `noauth` on all 19 focused actions and completes live workbook and PDF search/fetch flows in regular Chat. The same custom app class remains blocked by an incorrect expired-connection card in the Work surface before a tool call reaches the tunnel. Because the server is never invoked and the app succeeds in Chat, this is recorded as a ChatGPT Work host limitation; selecting **Chat** is the verified workaround.

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
- ChatGPT `fetch` now reads real `.xlsx` cells and formulas through the local Open XML helper, preserves worksheet context on both cold and indexed reads, and never depends on Graph text export for modern Office packages.
- The ChatGPT profile exposes 19 focused tools; standard `search` and `fetch` replace overlapping lookup and Office-read tools, preventing the redundant 12.9 MB workbook calls observed in the live host.
- The focused contract now has corrected fetch-first server instructions, per-tool selection cues and status text, plus an offline golden-prompt gate covering all 19 tools and six commonly confused tool pairs.
- ChatGPT search now supports a high-confidence stale-while-revalidate path with bounded background delta/query refresh; fetch validates unchanged indexed content by ETag/cTag and uses progressive 32 KiB previews with memory-backed 64 KiB continuation chunks.
- Bounded extraction covers common text/code formats, PDF, RTF, OpenDocument, EPUB, legacy Office, and common image OCR, with private temporary files, size limits, fixed extractor paths, timeouts, and cleanup.
- Upload, folder creation, rename, move, copy, file replacement/update, sharing-link create/revoke, recycle-bin delete/restore, and guarded permanent delete remain available on the focused ChatGPT surface with preview tokens and identity checks for risky mutations.

## Explicitly constrained coverage

- Business Graph Excel and organization-only sharing are mock-tested because the live account is personal.
- The isolated plus-address recipient was rejected by Graph; the harness verified that no grant existed, then completed live anonymous-link creation, audit, and exact revocation.
- Existing credentials, consent, and Keychain data remain untouched; forced login polling and Keychain deletion are excluded.
- Live recycle-bin restore remains excluded; native version restore passes live, while recycle-bin behavior passes mock and dry-run coverage.
- ChatGPT Work currently blocks this custom No Auth app with a false expired-connection card before tool dispatch. Regular Chat passes the same prompt and is the supported workaround pending a ChatGPT host correction.
- The current performance-hardening source commit is live on the Synology NAS and refreshed in ChatGPT; installation into the immutable Codex plugin cache remains pending and is outside this NAS/ChatGPT deployment.
- Synology administrative verification used HTTPS over the LAN fallback with insecure TLS certificate validation enabled after the Tailscale endpoint failed. The deployed service itself remains outbound-only, but DSM should be given a trusted certificate or certificate pinning to remove this operational warning.
