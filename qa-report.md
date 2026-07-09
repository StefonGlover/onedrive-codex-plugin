# QA Release Gate Report

Decision: Warn
Confidence: Medium
Date: 2026-07-09
Repository: <plugin-root>
Scope: full plugin repo plus live OneDrive beta
Mode: stabilization

## Executive Summary
- Local and installed-cache gates passed after fixes: Node syntax checks, setup script syntax, mocked Microsoft Graph regression suite, prepackage guard, and installed-cache drift guard.
- Live OneDrive beta passed: 75 checks, 0 failures, temporary folder cleanup confirmed, and local beta work cleaned.
- Search/indexing improvements now include batched metadata-cache writes, resumable delta `nextLink` cursors, progress milestones, duplicate folder-hint pruning, bounded concurrent fallback scans, exact-cache-hit confirmation, and faster local content-index matching.
- High-risk live delete/share/revoke/restore operations now require dry-run `previewToken` proof, and batch delete/move/revoke failures return explicit partial-remote-state warnings.
- Release decision remains Warn because full Microsoft 365/work-school live CRUD coverage was not available in this account; tenant-matrix script support was added and smoke-tested in read-only mode.

## Scope And Quality Oracle
- Sources used: README.md, skills/onedrive/SKILL.md, .codex-plugin/plugin.json, .mcp.json, .github/workflows/ci.yml, scripts/mock-graph-test.mjs, scripts/beta-test.mjs, scripts/prepackage-check.mjs, and mcp/server.mjs.
- The behavior oracle was the plugin README safety model, MCP tool schemas, mocked Graph regression suite, prepackage guard, and the live beta script.
- Existing uncommitted changes in mcp/server.mjs and scripts/mock-graph-test.mjs were present before this pass and were treated as active work, not reverted.

## Critical Paths Reviewed
- Auth and token handling: device-code flow, token check, logout memory-only path, and Keychain non-deletion behavior.
- File operations: list, scan, find, read, preview, download, upload, upload sessions, text writes, folder create, rename, move, copy, batch move, batch delete, and update_file checkout/commit.
- Sharing and permission safety: create link, invite, revoke, batch revoke, shared_by_me/public_links audit, dry-run defaults, confirmation requirements, expected identity checks, and audit logs.
- Release and packaging: CI workflow, manifest parsing, MCP schema inspection, prepackage file set checks, setup script syntax, and installed-cache drift guard behavior.

## Highest-Risk Findings
| Severity | Taxonomy | Area | Finding | Evidence | Residual Risk |
| --- | --- | --- | --- | --- | --- |
| Medium | confirmed finding | CI/release | `scripts/prepackage-check.mjs --installed` could false-pass without a path or when pointed at the source root. Fixed by requiring a path, rejecting source-root comparisons, and adding a CI negative guard. | scripts/prepackage-check.mjs:14, scripts/prepackage-check.mjs:131, .github/workflows/ci.yml:41 | Separate source-vs-installed cache comparison still requires a real installed path. |
| Medium | confirmed finding | security | Tool-visible Graph errors could expose recipient emails, object IDs, URLs, or bearer-looking strings. Fixed by redacting tool-facing and structured error fields. | mcp/server.mjs:2342, mcp/server.mjs:5319, scripts/mock-graph-test.mjs:1121 | Redaction is pattern-based; unusual provider-specific secrets may require future patterns. |
| Medium | confirmed finding | security | Copy monitor/resource URLs could expose query-bearing operation URLs, and trusted external copy monitor hosts did not require HTTPS. Fixed by enforcing HTTPS for external monitor hosts and stripping query strings from returned monitor references. | mcp/server.mjs:1957, mcp/server.mjs:4408, mcp/server.mjs:4468, scripts/mock-graph-test.mjs:978 | Returned host/path remains visible for supportability. |
| Medium | confirmed finding | reliability/performance | Recursive scan could follow unique empty pagination links without item progress. Fixed by adding a per-folder page cap and regression coverage. | mcp/server.mjs:2653, mcp/server.mjs:2664, scripts/mock-graph-test.mjs:797 | Extremely large folders remain bounded by configured caps and may need user tuning. |
| Medium | confirmed finding | safety/search | Ambiguous selector families, Keychain-token logout, update-file audit identity, `list_all` truncation, content-index refresh ordering, preview-token proof, and batch partial-state reporting had user-visible gaps. Fixed with validation and regression coverage. | mcp/server.mjs, scripts/mock-graph-test.mjs, scripts/beta-test.mjs, scripts/benchmark.mjs | Batch live operations remain inherently non-atomic in Microsoft Graph, but warnings and second-item-failure coverage are now explicit. |

## What Was Verified
- Mocked Microsoft Graph regression suite: 85 checks, 0 failures after fixes.
- Live OneDrive beta: 75 checks, 0 failures; temporary folder `Codex OneDrive Plugin Beta Test codex-beta-1783570238004-84416` was deleted and no longer resolved.
- Beta utility modes: `--doctor-only`, `--cleanup-stale` dry-run, and `--tenant-matrix=consumers` passed with 0 failures.
- Prepackage guard passed with 24 files checked.
- Installed-cache drift guard passed against `<installed-cache-path>`.
- Live read-only benchmark smoke: progress events emitted, cold find 9.7s with tight caps, cache refresh 5.9s for 20 items/5 folders, warm find 6.3s, content-index refresh reused 3 entries in 3ms, content search 1ms.
- Negative prepackage guard passed: missing `--installed` path now fails as expected.
- Git diff whitespace check passed.
- Setup script syntax check passed.
- Local CI-equivalent command bundle passed.

## What Was Not Verified
- Work/school live CRUD, macOS Keychain reset, Files.ReadWrite.All restore, and large tenant-scale Graph behavior were not exhaustively tested.
- CI was reviewed and simulated locally; GitHub Actions itself was not run remotely.

## Risk Acceptance
- Warn requires release-owner acceptance for remaining medium risks: full live CRUD was verified against the current personal OneDrive account only, and tenant-specific Microsoft Graph behavior can differ.

## Gate Results
| Gate | Status | Exit | Evidence |
| --- | --- | --- | --- |
| Syntax | pass | 0 | `node --check` for server and scripts; `zsh -n scripts/configure.zsh` |
| Mock tests | pass | 0 | `node scripts/mock-graph-test.mjs`, 85 checks, 0 failures |
| Live beta | pass | 0 | `node scripts/beta-test.mjs`, 75 checks, 0 failures, cleanup confirmed |
| Beta utility modes | pass | 0 | `--doctor-only`, `--cleanup-stale` dry-run, and `--tenant-matrix=consumers` passed |
| Prepackage | pass | 0 | `node scripts/prepackage-check.mjs`, 24 files checked |
| Installed drift | pass | 0 | `node scripts/prepackage-check.mjs --installed <installed-cache-path>`, 24 files checked |
| Negative prepackage | pass | 0 | `if node scripts/prepackage-check.mjs --installed; then exit 1; else exit 0; fi` |
| CI review | pass | n/a | CI now includes setup syntax and negative prepackage guard |

## Quality Model Coverage
| ISO/IEC 25010 Characteristic | Evidence | Residual Risk |
| --- | --- | --- |
| Functional suitability | Mock suite and live beta covered core file, search, indexing, sharing, and safety workflows. | Tenant/account-specific Graph differences remain. |
| Reliability | Retry paths, pagination cycles, scan page cap, cleanup, explicit partial batch failure reporting, and live beta cleanup were verified. | Batch operations can still partially complete after a mid-loop live failure, as disclosed in live responses. |
| Security | Error redaction, audit sanitization, upload-session trust, copy-monitor trust, dry-run defaults, and expected identity checks were verified. | Pattern-based redaction may need expansion for future provider messages. |
| Compatibility | Node 26 local run passed; CI targets Node 20. | Node 20 was not executed locally in this run. |
| Maintainability | CI-equivalent gates and focused regressions were added. | No package.json script entrypoint exists. |
| Portability/deployability | Prepackage guard, installed-cache drift check, and CI workflow review passed. | Packaging from a separate checkout should still run the installed drift guard before publishing. |

## Auto-Fix Summary
- Fixed confirmed CI/release false-pass in prepackage drift guard.
- Fixed confirmed security/privacy error redaction gaps.
- Fixed confirmed copy monitor URL trust and query exposure gaps.
- Fixed confirmed recursive scan empty pagination cap gap.
- Fixed ambiguous target selector validation, Keychain logout confirmation, update-file audit attribution, checkout local-path preflight ordering, `list_all` truncation reporting, `find_all` scan short-circuiting, content-result ranking balance, content-index refresh prioritization, preview-token proof, batch partial-state warnings, cache-refresh write batching, and delta `nextLink` continuation.
- Added regression coverage in the existing mocked Graph suite and CI workflow.

## Changes Made
- .github/workflows/ci.yml: adds setup script syntax check and negative `--installed` guard.
- scripts/prepackage-check.mjs: validates `--installed` path and rejects source-root drift comparisons.
- mcp/server.mjs: redacts tool-facing error fields, enforces trusted upload/copy URL handling, strips query strings from copy monitor output, caps scan pagination, rejects ambiguous selector families, guards Keychain logout, fixes `list_all` truncation, attributes update commits correctly, improves `find_all` scan behavior, prioritizes stale/missing content-index entries, batches metadata-cache writes, stores incomplete delta `nextLink`, adds cache-refresh progress, requires preview tokens for high-risk live operations, and returns explicit batch partial-state warnings.
- scripts/mock-graph-test.mjs: adds regressions for redaction, copy monitor safety, scan pagination cap, selector ambiguity, logout confirmation, list pagination, update-file audit/preflight, capped content-index refresh, batch revoke preflight, upload session trust, batch move cache updates, preview-token live calls, and second-item-failure paths for batch delete/move/revoke.
- scripts/beta-test.mjs: carries preview tokens through live high-risk operations, adds stale beta-folder cleanup mode, doctor-only mode, and tenant-matrix support.
- scripts/benchmark.mjs: exposes find/search scan timing and content-index refresh counters in benchmark summaries and emits progress events.

## Tests Added Or Updated
- Mock regression for redacted invite Graph error output and audit logs.
- Mock regression for copy monitor query stripping and insecure external monitor rejection.
- Mock regression for scan page cap on unique empty pagination chains.
- Mock regressions for ambiguous selector rejection, logout confirmation, consumed-pagination truncation, update-file sync-folder preflight, update-file audit identity, and capped content-index refresh ordering.
- Mock regression for pre-existing batch revoke, upload-session trust, and batch move cache behavior was preserved and verified.
- Mock regressions added for batch delete/move/revoke second-item live failures and partial-state warnings.
- Mock/live beta flows updated for preview-token proof on high-risk live operations.
- CI negative check for `scripts/prepackage-check.mjs --installed` without a path.

## Commands Run
| Command | Exit | Result | Notes |
| --- | --- | --- | --- |
| `python3 .../qa_repo_probe.py ... --json` | 0 | pass | Repo detected as git repo with initial dirty files. |
| `node --check mcp/server.mjs` | 0 | pass | Syntax. |
| `node --check scripts/mock-graph-test.mjs` | 0 | pass | Syntax. |
| `node --check scripts/prepackage-check.mjs` | 0 | pass | Syntax. |
| `node --check scripts/beta-test.mjs` | 0 | pass | Syntax. |
| `zsh -n scripts/configure.zsh` | 0 | pass | Setup script syntax. |
| `node scripts/mock-graph-test.mjs` | 0 | pass | 85 checks, 0 failures after fixes. |
| `node scripts/prepackage-check.mjs` | 0 | pass | 24 files checked. |
| `node scripts/prepackage-check.mjs --installed <installed-cache-path>` | 0 | pass | Installed cache matches source; 24 files checked. |
| `if node scripts/prepackage-check.mjs --installed; then exit 1; else exit 0; fi` | 0 | pass | Missing installed path fails as expected. |
| `git diff --check` | 0 | pass | No whitespace errors. |
| local CI-equivalent command bundle | 0 | pass | JSON manifests, syntax, mock tests, prepackage, negative prepackage. |
| `node scripts/beta-test.mjs` | 0 | pass | 75 checks, 0 failures; cleanup confirmed for `codex-beta-1783570238004-84416`. |
| `node scripts/beta-test.mjs --doctor-only` | 0 | pass | 5 checks, 0 failures. |
| `node scripts/beta-test.mjs --cleanup-stale --stale-days=9999 --cleanup-max-items=100 --cleanup-max-folders=20 --cleanup-max-results=5` | 0 | pass | Dry-run cleanup mode passed; 0 stale candidates. |
| `node scripts/beta-test.mjs --tenant-matrix=consumers` | 0 | pass | Read-only tenant matrix passed for consumers. |
| `node scripts/benchmark.mjs --query="project plan" --maxItems=20 --maxFolders=5 --maxFiles=3 --maxBytesPerFile=4096` | 0 | pass | Progress events emitted; bounded read-only benchmark passed. |

## Subagent Coverage
- CI/release/test reviewer found the prepackage `--installed` false-pass and CI setup-script coverage gap; both were fixed.
- MCP behavior reviewer identified error redaction, copy monitor URL safety, and scan pagination cap issues; all were fixed with regressions.

## Residual Risks
- Batch live operations preflight all entries but cannot be atomic against Microsoft Graph; mid-loop failures may leave partial remote state and now report this explicitly.
- Live beta covered the current signed-in personal account only; read-only tenant matrix support was smoke-tested for `consumers`, but not full Microsoft 365/work-school CRUD.

## Release, Rollback, And Monitoring Recommendations
- Release only with explicit acceptance of the residual medium risks above.
- Before packaging from a separate source checkout, run `scripts/prepackage-check.mjs --installed /absolute/path/to/installed/cache`.
- For rollout, monitor the local mutation audit log with `onedrive_audit_recent` after first real user workflows.
- Rollback is source-level revert of the four changed files plus reinstall/refresh of the plugin cache if already distributed.

## Recommended Next Actions
1. Run a second full live beta on a Microsoft 365/work-school tenant before broad release.
2. Exercise `--tenant-matrix-live` only with explicit test accounts for each tenant endpoint.
3. Track cache-refresh timings on larger drives now that write batching and delta continuation are implemented.
