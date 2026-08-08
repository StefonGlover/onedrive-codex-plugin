# OneDrive 0.6.4 End-to-End Beta Report

Decision: Pending — production release; ChatGPT Work end-to-end beta passed with explicit environment and recipient gaps
Date: 2026-08-08
Generated: 2026-08-08T01:09:58Z
Tested source base commit: `f7eca243b52c40a15845f4ce22dc76431551f186`
Plugin version: `0.6.4+codex.20260808010958`
Server version: `0.6.4+codex.20260808010958.chatgpt.3ade2b1c7f1f`
Tool contract: 84 exact tool names

## Outcome

The current candidate passed the focused 15-tool ChatGPT contract, OAuth and authorization compatibility, 187 mocked Microsoft Graph scenarios, genuine Office-package validation, a controlled live connector run, and a real ChatGPT Work UI mutation run. In ChatGPT itself, the model selected OneDrive and completed folder/text creation, guarded replacement, fetch verification, copy, move, permission inspection, owner-revoke refusal, anonymous-link create/revoke, recycle deletion, exact-item restore, and final recoverable cleanup.

The ChatGPT UI run exposed one additional contract bug: ChatGPT correctly carried `expectedETag` from a text-replacement preview, but the focused write schema rejected it. The schema and runtime now accept and enforce that optional revision guard. After deployment to `onedrive-chatgpt-nas:0.6.4-nas56-chatgpt-etag`, the same ChatGPT conversation retried the failed action successfully and finished the suite. The canary is healthy with zero restarts and zero failing streak; nas55 remains the rollback image.

## Live fixture and cleanup

The approved root was `Codex OneDrive Plugin Beta Test codex-beta-20260808t000304894z-8121`. All mutations stayed inside that boundary. Before cleanup it contained seven direct children and two nested text fixtures. The final root deletion used the guarded preview/confirmation flow and moved the folder to the OneDrive recycle bin. A direct path read then returned `itemNotFound`, confirming zero active fixture residue. No permanent deletion occurred and the recycle bin was not emptied.

The temporary anonymous view link was revoked and permissions returned to the single non-revocable owner grant before cleanup. No named invitation or email was sent because no distinct recipient was supplied.

The separate ChatGPT UI root `ChatGPT OneDrive Plugin Beta Test chatgpt-ui-20260808-01` was also recycled after its delete/restore test. ChatGPT verified the active path returned `itemNotFound`. Evidence: https://chatgpt.com/c/6a767e8f-0f4c-83ea-884d-9b406980876c

## Defects fixed

- Focused text-write schemas no longer expose overlapping destination branches; Work receives one unambiguous `remotePath` plus `content` contract.
- Guarded batch commits now report `partialMutationPossible: false` after a fully successful batch and reserve `true` for mixed success/failure state.
- Permission batch previews now fail closed for owner or inherited grants and never return a commit token for a non-revocable permission.
- Local Office-backed tests honor `ONEDRIVE_OFFICE_TEST_PYTHON`, then `ONEDRIVE_OFFICE_PYTHON`, then `python3`, instead of relying on a broken macOS system Python path.
- OAuth HTTP expectations now match the current 15-tool Work surface instead of the retired 21-tool profile.
- The versioned cache installer accepts the repository's `0.6.4` release line rather than rejecting it with a stale `0.6.1` guard.
- Installed-cache verification now evaluates audited sensitive source paths relative to the cache root, so the required OAuth facade token module is accepted consistently in source and installed snapshots.
- Focused text replacement now accepts and validates `expectedETag`, matching the preview/commit arguments naturally selected by ChatGPT Work.

The permission and partial-state regressions passed after nas55. The ChatGPT-specific replacement regression then passed in the original Work conversation after nas56 deployment, preserving the stable item ID, advancing the eTag, and verifying the expected 49-byte content fingerprint.

## Verification

- Mock Microsoft Graph: 187/187.
- Full MCP contract: 84 tools, 339,137 descriptor bytes.
- Focused ChatGPT contract: 15 tools, 26,526 bytes; 92.2% reduction.
- Focused OAuth contract: 15 tools, 28,656 bytes.
- Golden routing: 15/15 prompts and 8 ambiguity pairs.
- OAuth HTTP integration: pass.
- OAuth compatibility: 169/169.
- Semantic anchors: 6/6.
- Text patch safety: 6/6.
- Word local operation coverage: 21/21; genuine package reopened and rendered.
- Excel local operation coverage: 33/33; genuine package reopened and rendered.
- PowerPoint local operation coverage: 25/25; genuine package reopened and rendered.
- Live Office extraction and edits: DOCX, XLSX, PPTX, and PDF all passed; edited values were re-opened and verified.
- ChatGPT Work read-only UI probe: pass. Evidence: https://chatgpt.com/c/6a76730c-1764-83ea-98ff-be3a2cd18c48
- ChatGPT Work mutation/recovery UI suite: pass after one discovered-and-fixed schema regression. Evidence: https://chatgpt.com/c/6a767e8f-0f4c-83ea-884d-9b406980876c
- nas56 deployment: healthy, zero restarts, zero failing streak.
- Versioned Codex cache: `0.6.4+codex.20260808010958` installed alongside prior builds; 61-file source/cache parity and server SHA-256 passed.

## Improvement opportunities and remaining constraints

- One exact-filename search returned the correct restored file but took 49.7 seconds after filtering a stale cache item. The next improvement should cap exact-name fallback latency and expose per-stage timing in the user-visible result so indexed search, cache confirmation, and live scan can be distinguished.
- Named-recipient grant/revoke remains mock-tested only. Completing that live check requires a distinct recipient address and would send or grant external access.
- Business Graph Excel and organization-scoped sharing remain mock-tested because the connected OneDrive is personal.
- The extra ChatGPT attachment-to-OneDrive Office test is blocked by the Edge extension setting that permits local file URLs. Live Office upload/edit/export remains passed through the connector and local package suites.
- The NAS management connector's certificate-validation hardening remains separate infrastructure work.

The detailed capability matrix is in `work/qa-artifacts/capability-matrix-20260807.md`.
