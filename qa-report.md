# OneDrive 0.6.4 Release Gate Report

Decision: Pending — full cross-surface release; current regular-Chat deployment passes
Date: 2026-07-25
Generated: 2026-07-26T05:27:10Z
Tested source base commit: `146f724e5ee0b2c88690749e605449aefaa8e537`
Plugin version: `0.6.4+codex.20260729033800`
Server version: `0.6.1+codex.20260726012710.chatgpt.9956d1d6f66f`
Tool contract: 84 exact tool names

## Outcome

The `0.6.1` candidate adds a fail-closed public outer PKCE facade to the ChatGPT Work package mapping and delegated OAuth/OBO transport. It substitutes one-time mapped authorization codes, verifies S256 locally, and keeps provider refresh tokens in an encrypted owner-only vault behind rotating one-time handles. Both the 84-tool full profile and 21-tool ChatGPT profile pass the OAuth HTTP integration suite, including issuer-bound duplicate-key selection, bounded JWKS rotation refresh, authorized-client enforcement, RFC 9728 protected-resource discovery, HTTP 401 challenges, and non-relinking 503 provider/configuration failures.

The live NAS remains deliberately unchanged at the healthy no-auth rollback point `onedrive-chatgpt-nas:0.5.1-nas32` while the Entra registrations and ChatGPT developer app are configured. The new image and OAuth override will not replace that working runtime until discovery, OBO, and Work connection checks pass.

## Live regular-Chat coverage

The live Edge QA exercised:

- exact and descriptive search;
- folder and UTF-8 text-file creation;
- content replacement and verification;
- rename, move, and copy;
- permission inspection;
- anonymous view-only sharing-link creation and revocation;
- recycle-bin delete, deleted-item restore, and final recoverable cleanup.

The isolated fixture root `Codex OneDrive Edge QA 20260725-01` was moved to the recycle bin after testing. It remains recoverable. No permanent deletion occurred, the recycle bin was not emptied, no invitation or email was sent, and the temporary anonymous link was revoked. Permission count returned from 2 to 1 owner-only permission.

The live run exposed five ChatGPT tool-selection/schema problems. Focused descriptions now:

- require `content` rather than an invalid `text` field for `onedrive_write_text`;
- distinguish folder creation from sharing-link creation;
- constrain `onedrive_permissions` to item identity plus optional format;
- reserve `onedrive_delete` for active-item recycle intent;
- reserve `onedrive_restore_deleted` for one restore while an item is recycled.

The post-deployment read-only smoke then exposed a Microsoft Graph discovery/indexing failure: direct paths resolved both known PDFs, but indexed search and the exact-file opener returned no matches. `onedrive_open_files` now validates cached exact identities and, only when no exact indexed hit exists, performs a bounded read-only live folder scan capped at 2,000 items, 300 folders, depth 20, 10 results, and concurrency 3.

After deploying `nas32` and refreshing ChatGPT, a single `onedrive_open_files` call in regular Chat returned both targets as `found`:

| File | Item ID | Path | Permissions |
| --- | --- | --- | --- |
| `invoice-3095.pdf` | `B8C89DB91F19C763!s2c28febab675475d813f0cab07fd4e36` | `Family Space/Documents/Home/invoice-3095.pdf` | 1, owner-only |
| `2026 Electrical Report.pdf` | `B8C89DB91F19C763!s66748b567d5745b48e96db533485d4ff` | `Family Space/Documents/Home/Electrical/2026 Electrical Report.pdf` | 1, owner-only |

The final smoke reported no lookup, schema, host, authentication, or permission-inspection errors and made no changes. Evidence is in the regular Chat thread `https://chatgpt.com/c/6a6538f3-d740-83ea-8719-af1308a75b70`.

## Verification

- Node syntax/self-check: pass.
- Full contract: 84 tools, 338,721 bytes.
- ChatGPT contract: 21 tools, 34,134 bytes without OAuth and 37,830 bytes with OAuth, 89.9% reduction from the full descriptor.
- Server instructions: 1,316 bytes.
- Golden prompts: 21/21 with 11 ambiguity pairs.
- Mock Microsoft Graph: 176/176.
- Semantic anchors: 6/6.
- Text patch safety: 6/6.
- OAuth HTTP integration: pass in full (84 tools) and ChatGPT (21 tools) profiles.
- Prepackage inventory: 59 files.
- ChatGPT refreshed action descriptions: pass, including all five routing guards and the exact-file fallback.
- NAS image, container health, restart count, staged source, and rollback: pass.
- Final regular-Chat exact-file and permissions regression: pass.

## Deployment notes and constraints

- A stale local launch-agent tunnel was competing with the NAS endpoint. `com.stefonglover.onedrive-chatgpt-tunnel.plist` was booted out, and no local OneDrive tunnel/server process remains. The NAS is now the sole live endpoint.
- ChatGPT Work is pending live Entra registration, developer-app OAuth configuration, NAS OAuth cutover, and a new Work-chat validation. Its package mapping and transport are implemented and mock-tested.
- Named-recipient invitation was not exercised live to avoid sending email or granting access to another person. Its guarded behavior remains mock-tested.
- Business-only Graph Excel and organization-only sharing remain mock-tested because the connected account is personal.
- The deployed developer app currently uses the encrypted server-side vault with ChatGPT `No Auth`; it remains the rollback path until production Entra/OAuth validation completes.
- The existing NAS management connection disables TLS certificate verification. This did not affect runtime health, but certificate validation should be hardened separately.
