# OneDrive 0.5.1 Release Gate Report

Decision: Pending — Entra registration, NAS OAuth rollout, and ChatGPT Work validation
Date: 2026-07-21
Generated: 2026-07-21T19:59:32Z
Tested source base commit: `3bd2f3d6617a74843dd7a953ff51305e74ac6f0d`
Plugin version: `0.5.1+codex.20260721181802`
Tool contract: 84 exact tool names

## Current outcome

The real-productivity beta passes on the canonical ChatGPT developer app **OneDrive** (`asdk_app_6a5e2416985481918d0f6c68785da2c4`). ChatGPT metadata version `dev-22` exposes the focused 21-tool No Auth surface, retains the stored local OneDrive logo, and runs the final source contract `0.5.1+codex.20260721181802.chatgpt.2d26f01bba1b`.

The final NAS runtime is `onedrive-chatgpt-nas:0.5.1-nas27`, built as image `81bd2d14c6a6` from the fresh versioned source directory `/volume1/docker/onedrive-chatgpt/app-0.5.1-nas27-20260721`. Container Manager recreated and started the project with exit code 0. The prior `nas26` image is retained as rollback, and `compose.nas25.rollback.yaml` preserves the earlier manifest.

The focused ChatGPT contract is 21 tools and 40,536 bytes, 88.0% smaller than the unchanged 84-tool, 338,721-byte full contract. The 1,316-byte server instructions and focused descriptors now prefer user-visible OneDrive paths over opaque item IDs. The ChatGPT copy schema no longer advertises `waitForCompletion` or `timeoutSeconds`, avoiding unnecessary inline Graph polling while preserving separate verification.

Search now handles a whole multi-document intent in one pass, merges fallback results instead of replacing good matches, ranks requested document kinds and years, and verifies uncovered concept domains only. RFC 822 `.eml` extraction now returns bounded headers, readable text bodies, and attachment inventories. Modern Office, PDF, text/code, RTF, OpenDocument, EPUB, legacy Office, and common image handling remain available.

The live productivity beta stayed in one ChatGPT thread: `https://chatgpt.com/c/6a5f7eb8-afb0-83ea-a34c-9c2b5cb77ea8`.

- One multi-document search correctly returned `invoice-3095.pdf` and `2026 Electrical Report.pdf` in about 32 seconds.
- Explicit `.eml` extraction returned subject, sender, date, readable body, and all named attachments in 15 seconds.
- Disposable folder/file creation completed in 25 seconds.
- Copy, rename, and dependent patch operations completed in sequence with fresh identities and no stale retry; the longer host flow took 1 minute 36 seconds.
- An anonymous view-only link was created in 11 seconds, then revoked in 28 seconds; the file was owner-only afterward.
- Both disposable QA folders were moved to the recycle bin, no item was permanently deleted, no anonymous link remains, and active OneDrive searches no longer find the fixtures.
- The final `dev-22` read-only smoke in the same thread again returned exactly `invoice-3095.pdf` and `2026 Electrical Report.pdf` with no mutation.

DSM staging cleanup removed the uploaded archive and moved the obsolete `app-0.5.1-nas17-20260721` directory to the NAS recycle bin. Exact local temporary archives, manifest copies, and deployment screenshots were removed. The fresh `nas27` source directory, active manifest, encrypted runtime/data, `nas26` rollback image, and rollback manifest remain.

## Verification

- Node syntax check: pass.
- Common extraction fixtures: RTF, OpenDocument, EPUB, and email pass.
- Full contract: 84 tools, 338,721 bytes.
- ChatGPT contract: 21 tools, 40,536 bytes, 88.0% reduction.
- Golden prompts: 21/21, with eight ambiguity pairs.
- Mock Microsoft Graph: 174/174.
- Semantic anchors: 6/6.
- Text patch preservation and safety: 6/6.
- Office Open XML operations: 79 total (21 Word, 33 Excel, 25 PowerPoint).
- Storage-root permissions: pass.
- Whitespace: pass.
- Live NAS image/tag, ChatGPT metadata, logo, focused tool count, path-first schemas, and final same-thread search: pass.

## Fixes validated

- `.eml` files are extracted locally instead of being treated as opaque text-export failures.
- Search understands subtle service-language aliases, document-kind intent, multi-domain requests, and year/report recency without discarding strong initial matches.
- Search and fetch return user-visible paths so ChatGPT can keep mutation flows readable and avoid mistaking opaque item IDs for credentials.
- Dependent mutations are serialized and refetch current item identity before the next preview or write.
- Copy returns its accepted asynchronous result promptly; ChatGPT verifies the destination separately.
- Folder creation is a direct conflict-safe create, while risky mutations retain preview tokens, expected identity, confirmation, and audit protections.
- Upload, folder creation, rename, move, copy, replacement/update, sharing-link creation/revocation, recycle-bin delete/restore, guarded permanent deletion, and permission inspection remain available on the focused surface.
- The stored OneDrive logo remains correct because the canonical developer-mode app was refreshed in place rather than recreated.

## Explicitly constrained coverage

- Business-only Graph Excel and organization-only sharing remain mock-tested because the connected account is personal.
- Live recycle-bin restore is excluded; recycle, restore, and permanent-delete safety contracts are covered by mocks and previews.
- ChatGPT Work still shows an incorrect expired-connection/authentication card before a tool call reaches this custom No Auth app. Regular Chat is the verified workaround pending a ChatGPT host correction.
- Installation into the immutable Codex plugin cache remains pending and does not affect the deployed NAS/ChatGPT app.
- Entra protected-resource discovery and OAuth integration pass locally but still require the production Entra registrations and NAS OAuth rollout before this release gate can become Pass.
