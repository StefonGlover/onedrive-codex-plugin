# QA Release Gate Report

Decision: Warn
Confidence: Medium
Date: 2026-07-07
Repository: <plugin-root>
Scope: full plugin repo plus live OneDrive beta
Mode: stabilization

## Executive Summary
- Full local CI-equivalent gates passed after fixes: JSON manifests, Node syntax checks, setup script syntax, mocked Microsoft Graph regression suite, prepackage guard, negative prepackage guard, and diff whitespace.
- Live OneDrive beta passed after fixes: 75 checks, 0 failures, temporary folder cleanup confirmed, and local beta work cleaned.
- Release decision is Warn, not Pass, because medium residual risks remain around ambiguous selectors, non-atomic batch mutations, and single-account live environment coverage.

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

## What Was Verified
- Mocked Microsoft Graph regression suite: 70 checks, 0 failures after fixes.
- Live OneDrive beta: 75 checks, 0 failures after fixes; temporary folder `Codex OneDrive Plugin Beta Test codex-beta-1783466428468-59354` was deleted and no longer resolved.
- Prepackage guard passed with 23 files checked.
- Negative prepackage guard passed: missing `--installed` path now fails as expected.
- Git diff whitespace check passed.
- Setup script syntax check passed.
- Local CI-equivalent command bundle passed.

## What Was Not Verified
- A real separate installed-cache drift comparison was not run because no distinct installed target path was supplied.
- Cross-account, work/school tenant, macOS Keychain reset, Files.ReadWrite.All restore, and large tenant-scale Graph behavior were not exhaustively tested.
- CI was reviewed and simulated locally; GitHub Actions itself was not run remotely.

## Risk Acceptance
- Warn requires release-owner acceptance for remaining medium risks: ambiguous selector precedence, batch operations with possible partial remote side effects after preflight, and single-account live beta coverage.

## Gate Results
| Gate | Status | Exit | Evidence |
| --- | --- | --- | --- |
| Syntax | pass | 0 | `node --check` for server and scripts; `zsh -n scripts/configure.zsh` |
| Mock tests | pass | 0 | `node scripts/mock-graph-test.mjs`, 70 checks, 0 failures |
| Live beta | pass | 0 | `node scripts/beta-test.mjs`, 75 checks, 0 failures, cleanup confirmed |
| Prepackage | pass | 0 | `node scripts/prepackage-check.mjs`, 23 files checked |
| Negative prepackage | pass | 0 | `if node scripts/prepackage-check.mjs --installed; then exit 1; else exit 0; fi` |
| CI review | pass | n/a | CI now includes setup syntax and negative prepackage guard |

## Quality Model Coverage
| ISO/IEC 25010 Characteristic | Evidence | Residual Risk |
| --- | --- | --- |
| Functional suitability | Mock suite and live beta covered core file, search, sharing, and safety workflows. | Tenant/account-specific Graph differences remain. |
| Reliability | Retry paths, pagination cycles, scan page cap, cleanup, and live beta cleanup were verified. | Batch operations can still partially complete after a mid-loop live failure. |
| Security | Error redaction, audit sanitization, upload-session trust, copy-monitor trust, dry-run defaults, and expected identity checks were verified. | Pattern-based redaction may need expansion for future provider messages. |
| Compatibility | Node 26 local run passed; CI targets Node 20. | Node 20 was not executed locally in this run. |
| Maintainability | CI-equivalent gates and focused regressions were added. | No package.json script entrypoint exists. |
| Portability/deployability | Prepackage guard passed and CI workflow was strengthened. | Separate installed-cache drift check still requires a distinct target path. |

## Auto-Fix Summary
- Fixed confirmed CI/release false-pass in prepackage drift guard.
- Fixed confirmed security/privacy error redaction gaps.
- Fixed confirmed copy monitor URL trust and query exposure gaps.
- Fixed confirmed recursive scan empty pagination cap gap.
- Added regression coverage in the existing mocked Graph suite and CI workflow.

## Changes Made
- .github/workflows/ci.yml: adds setup script syntax check and negative `--installed` guard.
- scripts/prepackage-check.mjs: validates `--installed` path and rejects source-root drift comparisons.
- mcp/server.mjs: redacts tool-facing error fields, enforces trusted upload/copy URL handling, strips query strings from copy monitor output, and caps scan pagination.
- scripts/mock-graph-test.mjs: adds regressions for redaction, copy monitor safety, scan pagination cap, batch revoke preflight, upload session trust, and batch move cache updates.

## Tests Added Or Updated
- Mock regression for redacted invite Graph error output and audit logs.
- Mock regression for copy monitor query stripping and insecure external monitor rejection.
- Mock regression for scan page cap on unique empty pagination chains.
- Mock regression for pre-existing batch revoke, upload-session trust, and batch move cache behavior was preserved and verified.
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
| `node scripts/mock-graph-test.mjs` | 0 | pass | 70 checks, 0 failures after fixes. |
| `node scripts/prepackage-check.mjs` | 0 | pass | 23 files checked. |
| `if node scripts/prepackage-check.mjs --installed; then exit 1; else exit 0; fi` | 0 | pass | Missing installed path fails as expected. |
| `git diff --check` | 0 | pass | No whitespace errors. |
| local CI-equivalent command bundle | 0 | pass | JSON manifests, syntax, mock tests, prepackage, negative prepackage. |
| `node scripts/beta-test.mjs` | 0 | pass | 75 checks, 0 failures; cleanup confirmed. |

## Subagent Coverage
- CI/release/test reviewer found the prepackage `--installed` false-pass and CI setup-script coverage gap; both were fixed.
- MCP behavior reviewer identified error redaction, copy monitor URL safety, and scan pagination cap issues; all were fixed with regressions.

## Residual Risks
- Ambiguous target selectors can still be supplied together; current precedence should be documented or rejected in a future change.
- Batch live operations preflight all entries but cannot be atomic against Microsoft Graph; mid-loop failures may leave partial remote state.
- Live beta covered the current signed-in account only, not a full matrix of personal and Microsoft 365 tenants.
- Release artifact comparison against a separate installed cache path was not run.

## Release, Rollback, And Monitoring Recommendations
- Release only with explicit acceptance of the residual medium risks above.
- Before packaging from a separate source checkout, run `scripts/prepackage-check.mjs --installed /absolute/path/to/installed/cache`.
- For rollout, monitor the local mutation audit log with `onedrive_audit_recent` after first real user workflows.
- Rollback is source-level revert of the four changed files plus reinstall/refresh of the plugin cache if already distributed.

## Recommended Next Actions
1. Decide whether to reject ambiguous target selector combinations in schemas/validation.
2. Add second-item-fails live-loop mock tests for batch delete, move, and revoke partial-state reporting.
3. Run installed-cache drift comparison against a distinct packaged/installed target before publishing.
