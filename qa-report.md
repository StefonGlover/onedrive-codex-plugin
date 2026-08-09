# OneDrive 0.6.4 End-to-End Beta Report

Decision: Pending — broader release sign-off; NAS deployment and ChatGPT UI betas pass with no plugin defects open
Date: 2026-08-09
Generated: 2026-08-09T07:13:35Z
Tested source base commit: `abd9072ab933dc1dac1cd6341c9f36c092a5d55f`
Plugin version: `0.6.4+codex.20260809062901`
Focused source server version: `0.6.4+codex.20260809062901.chatgpt.3ade2b1c7f1f`
Tool contract: 84 exact tool names; 15 focused ChatGPT tools

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
