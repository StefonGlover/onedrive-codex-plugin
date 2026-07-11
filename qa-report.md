# QA Release Gate Report

Decision: Pass for personal OneDrive beta; Warn for broad tenant rollout
Confidence: High for tested scope
Date: 2026-07-11
Repository: <plugin-root>
Scope: OneDrive plugin source, installed cache 0.1.0+codex.20260711152309, mock Microsoft Graph suite, CLI guards, performance benchmarks, and personal-account live beta

## Executive Summary

- All discovered and reproduced defects from this beta run were fixed or turned into explicit safe behavior.
- Official mocked Microsoft Graph suite passes on Node 26: 110 checks, 0 failures.
- Independent fixed-hash verifier passes: 10/10 adversarial categories, 0 failures, no residue.
- Source and installed-build live personal OneDrive betas each passed 81/81, covering all 58 tools with remote and isolated-local cleanup confirmed.
- Installed-build read-only tenant matrix passed for common, consumers, and organizations.
- Broad tenant rollout remains Warn because full work/school Microsoft 365 CRUD was not available in this account; read-only tenant health support is covered by the beta script.

## Issues Fixed

| Area | Fix |
| --- | --- |
| Search speed | `onedrive_find` now runs the canonical Graph query first, stops normal expansion after confident matches, executes `find_all` exhaustively, batches search cache writes, and treats `graphSearchCalls` as actual pages fetched. |
| Fallback scans | Completed hinted subtrees are excluded from the final root fallback without excluding truncated subtrees; controlled cold-miss time improved 12%, from 22.38s to 19.69s, and `foldersSkipped` makes the optimization observable. |
| Cache correctness | Metadata and content-index files now use compact atomic writes, process/interprocess locks, generation-aware reloads, and cache version 3 cursor migration. |
| Delta sync | Delta cursor persistence is scoped to delta-origin cursors and matching roots; pathless delta records hydrate through parent IDs where possible and unresolved paths are reported. |
| Content index | cTag/eTag handling preserves metadata-only renames and invalidates conservatively when content freshness is uncertain. |
| Local outputs | Concurrent downloads, exports, update checkouts, backups, and audit exports reserve unique paths or fail safely without clobbering. |
| Beta isolation | Every beta run uses unique temporary storage/cache roots, validates all 58 tools, exercises the content-index lifecycle and batch permission revoke, and cleans isolated state even after failures. |
| Office downloads | Excel, Word, and PowerPoint helpers reuse resolved metadata, eliminating one Graph request per download. |
| Packaging | Replaced the stale 55-tool installed cache with a cache-busted 58-tool build and verified byte-for-byte source/install parity. |
| Audit/auth | Per-tool request IDs, failed-call `localWarnings`, audit locking, and device-login supersession are isolated under concurrency. |
| Utilities | Beta cleanup is bounded and timestamp-safe; invalid CLI numbers/booleans fail early; benchmark tool errors exit nonzero; prepackage installed-cache checks reject bad paths. |

## Verification

| Gate | Result | Evidence |
| --- | --- | --- |
| Syntax | Pass | `node --check` for server and scripts |
| Whitespace | Pass | `git diff --check` |
| Prepackage | Pass | `node scripts/prepackage-check.mjs`, 24 files checked |
| Installed drift | Pass | source vs installed cache 0.1.0+codex.20260711152309, 24 files checked |
| Beta self-check | Pass | 17/17 utility checks |
| Benchmark self-check | Pass | 7/7 utility checks |
| Mock suite, Node 26 | Pass | 110/110 |
| Independent verifier | Pass | 10/10 adversarial categories; hashes stable; no residue |
| Negative CLI guards | Pass | invalid stale days, cleanup page size, retry attempts, booleans, and installed-cache paths rejected |
| Source live beta | Pass | 81/81; 58/58 tools; isolated remote folder and local state deleted |
| Installed live beta | Pass | 81/81; 58/58 tools; isolated remote folder and local state deleted |
| Installed tenant matrix | Pass | common 5/5, consumers 5/5, organizations 5/5 |
| Installed cleanup dry-run | Pass | stale-days=0; no beta folder candidates |
| Performance benchmark | Pass | completed hinted-subtree dedup improved controlled cold miss 12%; exact canonical hits remain one Graph search |
| Shared cache repair | Pass | full root scan rebuilt 2,147 metadata entries and stored a fresh delta cursor after pre-fix beta state loss |
| Managed storage modes | Pass | directories 0700; cache/index/audit files 0600 |

## Independent Verifier Coverage

- Same-process and cross-process metadata/content-index retention.
- Atomic JSON visibility under concurrent reads.
- Same-name download/export reservation and checkout collision refusal.
- UUID default audit exports.
- Root and scoped delta hydration with no cursor poisoning.
- cTag rename reuse and omitted-cTag invalidation.
- Exhaustive `find_all` term execution and paginated search-call accounting.
- Failed mutation `localWarnings`, concurrent request-ID isolation, and device-login supersession.

## Residual Risk

- Microsoft Graph batch live operations remain non-atomic by design; the plugin preflights, warns, and reports partial-state failures.
- Full live CRUD was verified only against the available personal OneDrive account. Work/school tenant CRUD should be rerun before broad release.
- Pattern-based redaction may need future expansion for new provider error shapes.

## Install Result

- Installed cache: `<personal-cache>/onedrive/0.1.0+codex.20260711152309`
- Source-vs-installed drift check: pass, 24 files checked.
- Final stale cleanup discovery: pass, 0 beta folder candidates.
- Local test residue: cleaned from source and installed cache.
