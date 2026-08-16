# OneDrive Codex Plugin

Local Codex plugin for OneDrive file operations through Microsoft Graph.

Release `0.6.4+codex.20260816170509` adds five remotely importable workflow skills, clearer focused Office schemas, explicit personal-account responses for enterprise-only reads, a subject-scoped 60-second folder-list snapshot with mutation invalidation, and NAS startup ownership for private materialized downloads/previews. It retains filename-only visible hyperlinks, direct link opening, the 19-tool ChatGPT surface, the full 84-tool trusted-local profile, account-and-drive-isolated state, bounded Office inspection and review, guarded version restore, enterprise exact-drive reads, and prior nested-folder performance improvements.

ChatGPT surface note (updated 2026-08-13): the focused 19-tool surface covers discovery, Office inspection/editing/review, materialized downloads and visual previews, recent/version history, enterprise drive libraries, CRUD, sharing, and permissions. Redundant standalone tools remain available only in the full profile; a focused-profile `tools/call` now rejects every unadvertised tool instead of treating a hidden handler as executable compatibility surface. Work requires the delegated OAuth deployment and public compatibility origin described below. Direct Entra v2 endpoints are not sufficient because ChatGPT correctly sends the MCP `resource` parameter and Entra v2 rejects that parameter.

This is an unofficial integration and is not affiliated with, endorsed by, or sponsored by Microsoft.

All plugin icon surfaces use `assets/chatgpt-icon.png`, the square 256×256 OneDrive image derived from the supplied artwork. The MCP initialization metadata also embeds that PNG as a standards-based `serverInfo.icons` data URI. ChatGPT stores the directory-listing logo separately, so upload this PNG in the **Icon** field during the app's initial developer-mode setup.

The plugin is remote-first: it uses Microsoft Graph rather than the laptop's local OneDrive sync folder. Upload and download tools refuse local OneDrive sync-folder paths by default unless `allowLocalOneDriveSyncPath: true` is explicitly provided.

## Setup

Local/Codex and legacy No Auth deployments use Microsoft identity platform device-code login. On macOS they store refresh tokens in Keychain. On Linux/NAS they use an AES-256-GCM encrypted file protected by a separate owner-only key file. ChatGPT Work OAuth requests carry a short-lived Entra token for the MCP API; the server validates it and exchanges it for a short-lived Graph token with the on-behalf-of flow. The public OAuth facade never stores passwords or plaintext refresh tokens: it stores the upstream refresh token only as authenticated ciphertext and gives ChatGPT a hashed, rotating one-time handle.

1. Create or choose a Microsoft Entra app registration.
2. Enable public client flows for the app.
3. Add delegated Microsoft Graph permissions for `User.Read`, `Files.ReadWrite`, and `offline_access`.
   - Optional: add `Files.ReadWrite.All` if you want to attempt live restore of deleted personal OneDrive items with `onedrive_restore_deleted`.
4. Run the local setup script:

```bash
scripts/configure.zsh
```

5. Start a fresh Codex thread after installing or refreshing the plugin.
6. Ask Codex to call `onedrive_config` with `checkToken: true`. If no reusable credential exists, call `onedrive_auth_device_start`, open the returned verification URL, enter the returned user code, then call `onedrive_auth_device_poll`.

After the first successful login, the refresh token is reused from the configured secure authentication store. `onedrive_auth_device_start` checks that credential before contacting Microsoft's device-code endpoint and returns `alreadyAuthenticated: true` without generating a code when authentication is healthy. Use `forceReauth: true` only for an intentional account switch, consent repair, or explicit sign-in reset. Temporary token-check network failures do not trigger a new login flow.

## ChatGPT Developer mode

The MCP server can also run as a private ChatGPT developer-mode app through OpenAI's [Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels). Every tool descriptor includes the required read-only, open-world, and destructive impact annotations.

1. Enable **Developer mode** in ChatGPT under **Settings → Security and login**.
2. Create a tunnel in [OpenAI Platform tunnel settings](https://platform.openai.com/settings/organization/tunnels), and associate both the Platform organization and the target ChatGPT workspace.
3. Install the latest official `tunnel-client` release and create a runtime API key whose principal has Tunnels **Read + Use**.
4. Save that key outside the plugin source tree at `~/.config/tunnel-client/onedrive-chatgpt.env` as `OPENAI_API_KEY` or `CONTROL_PLANE_API_KEY`, and restrict the file to owner-only access with `chmod 600`. Keeping the runtime credential outside the plugin prevents it from entering a packaged cache version.
5. Create the local stdio profile. Quote the server path inside `--mcp-command`; unquoted paths containing spaces are split into invalid Node arguments.

```bash
onedrive_plugin_root="$(pwd)"
tunnel-client init \
  --sample sample_mcp_stdio_local \
  --profile onedrive-chatgpt \
  --tunnel-id tunnel_... \
  --mcp-command "node \"$onedrive_plugin_root/mcp/server.mjs\""
```

6. Validate the stopped profile with `tunnel-client doctor --profile onedrive-chatgpt --explain`, then run it with `node scripts/run-chatgpt-tunnel.mjs`. The launcher reads the private tunnel-client env file without printing the key and locks hosted use to `ONEDRIVE_TOOL_PROFILE=chatgpt` (the safe default), a focused 19-tool contract with compact Office schemas for lower ChatGPT selection latency. It combines independent reads, guarded commits, targeted Office inspection/review, materialized downloads, and visual QA without exposing setup or maintenance handlers. Unadvertised direct calls fail closed in this profile. Setup, diagnostics, cache/index maintenance, specialized reporting, bulk automation, and local-path tools remain available only with the trusted local stdio `full` profile; Streamable HTTP and ChatGPT tunnel launchers reject `full` at startup. Every focused tool reports a short in-progress/completed status. Set `ONEDRIVE_TUNNEL_ENV_FILE` only when you intentionally use a different credential-file path.
7. In ChatGPT **Settings → Plugins**, create a developer-mode plugin, choose **Tunnel**, select the tunnel, choose **No Auth**, create the plugin, and connect it.

Keep `tunnel-client` running for connector discovery and every OneDrive tool call from ChatGPT.

### ChatGPT Work OAuth

Work uses an Entra-protected MCP API and Microsoft Graph on-behalf-of (OBO) exchange. This requires two Entra app registrations:

1. **OneDrive MCP API** — the existing OneDrive client registration can be reused.
   - Keep its delegated Graph permissions: `User.Read` and `Files.ReadWrite` (plus any optional permission you already use).
   - Under **Expose an API**, set the Application ID URI to `api://<MCP_API_CLIENT_ID>` and add a delegated `access_as_user` scope that users and admins can consent to.
   - Add a confidential-client secret or certificate. The current implementation accepts the secret through `ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET_FILE`; never commit it.
   - Grant the existing user/admin consent for the Graph delegated permissions so OBO does not fail with `interaction_required`.
2. **ChatGPT OneDrive Client** — create a separate confidential web-client registration used by the facade for upstream Entra exchange. Prefer the outer public-PKCE mode, where ChatGPT never receives this registration's secret. The explicitly enabled confidential compatibility mode described below instead gives the same client secret to ChatGPT for `client_secret_post` token authentication.
   - Support the same account population as the API registration (`consumers` for this personal OneDrive, or the appropriate tenant for work/school accounts).
   - Add delegated permission to `api://<MCP_API_CLIENT_ID>/access_as_user`.
   - In the API registration's **Authorized client applications**, pre-authorize this client ID for `access_as_user` when your tenant policy permits it.
   - In preferred public-PKCE mode, add the facade callback `<ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER>/callback` as the Entra Web redirect URI. In confidential compatibility mode, also register ChatGPT's exact `https://chatgpt.com/connector/oauth/<callback_id>` Web redirect because Microsoft returns the code there directly.

Configure the NAS with `deploy/synology/compose.oauth.example.yaml`. Put the MCP API registration's secret in `deploy/synology/runtime/oauth-api-client.secret` and the facade's private upstream Entra client secret in `deploy/synology/runtime/oauth-chatgpt-client.secret`; keep both files mode `0600`. Set `ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS` to the **ChatGPT OneDrive Client** application ID so a valid API token issued to some other Entra client is still rejected.

Keep the resource, authorization, and validation identities separate:

- The Entra API resource is `api://<MCP_API_CLIENT_ID>`. It defines `ONEDRIVE_MCP_OAUTH_API_RESOURCE`, the delegated `ONEDRIVE_MCP_OAUTH_API_SCOPE`, and the token audience checked through `ONEDRIVE_MCP_OAUTH_AUDIENCE`.
- The protected MCP resource is the exact externally visible HTTPS tunnel MCP endpoint. Put it in `ONEDRIVE_MCP_PROTECTED_RESOURCE`, and put the exact public protected-resource metadata URL advertised for that tunnel in `ONEDRIVE_MCP_RESOURCE_METADATA_URL`. Never put an `api://` identifier or `127.0.0.1` URL in either public-MCP setting.
- The advertised authorization server is the compatibility service's public HTTPS origin. Put it in both `ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER` and `ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER`.
- The inbound-token validation authority remains Microsoft Entra in `ONEDRIVE_MCP_OAUTH_AUTHORITY`; the OBO token endpoint is derived from the already validated `tid` claim unless `ONEDRIVE_MCP_OAUTH_OBO_TOKEN_ENDPOINT` is explicitly pinned.

The MCP Streamable HTTP process stays on loopback behind Secure MCP Tunnel. `ONEDRIVE_MCP_HTTP_HOST` accepts only `127.0.0.1`, `::1`, or `localhost`; an unset or blank value uses `127.0.0.1`, while wildcard, LAN, and other hostname bindings are rejected before the server starts. The separate compatibility service is mapped to NAS loopback port `3010` and must be exposed through a trusted public HTTPS reverse proxy such as Tailscale Funnel. Its public surface is limited to OAuth metadata, `/authorize`, `/callback`, `/token`, and `/healthz`. Because Funnel uses a shared loopback peer, no limiter trusts the socket address or forwarded headers. In outer-public mode the facade requires S256, rate-binds the validated state/PKCE pair, applies route-wide count-and-byte admission budgets, forces the Entra callback to query mode, rejects callback POST bodies, and returns one-time mapped codes plus rotating refresh handles. Token POSTs are protected by bounded pre-parse concurrency, strict body/time limits, and a capped LRU keyed only by a domain-separated hash of the validated facade code or handle. Outer state is capped at 512 UTF-8 bytes and provider codes at 3 KiB. The mode-`0600` encrypted store reserves independent 320 KiB refresh and 640 KiB ephemeral partitions inside its 1 MiB total cap; plaintext codes, refresh tokens, client secrets, and request values are excluded from logs and metadata. Its encryption key is domain-derived from the stable owner-only auth-vault key, so routine OAuth client-secret rotation preserves active sessions; rotating the auth-vault key intentionally requires a clean reconnect. Before cutover, verify the public compatibility metadata and run `tunnel-client doctor --profile onedrive-chatgpt --explain`.

The plugin's `.app.json` maps the package to the exact registered ChatGPT connection ID copied from its developer-mode URL. The ID must use the `plugin_asdk_app_...` form; it is not an Entra application ID and must not be replaced by a bare `asdk_app_...` value.

For ChatGPT Work, advertise Client ID Metadata Document support and use the public-client authorization-code flow with token authentication method `none` and mandatory S256 PKCE. Set `ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD=true`, `ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR=false`, and both no-PKCE compatibility switches to `false`. Register the separate facade callback `<ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER>/callback` in Entra while keeping ChatGPT's fixed callback in `ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI`. Request:

```text
api://<MCP_API_CLIENT_ID>/access_as_user offline_access
```

Do not add `openid` or `profile`: the compatibility tier is deliberately OAuth-only. It forwards Entra access tokens but never exposes Entra refresh tokens; refresh uses encrypted, rotating facade handles instead.

When ChatGPT Work rewrites a Secure MCP Tunnel resource to its internal
tunnel-gateway identifier, put only that exact same-tunnel value in
`ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES`. The facade accepts aliases
only from the bounded OpenAI gateway host pattern and only when the tunnel ID
matches the canonical `https://api.openai.com/v1/mcp/tunnel_...` resource.
Metadata stays canonical, while the exact validated resource from ChatGPT
propagates to the access-token audience. Configure the same value in
`ONEDRIVE_MCP_PROTECTED_RESOURCE_ALIASES` so the MCP accepts that equivalent,
same-tunnel audience.

The DCR implementation stays covered by compatibility tests but is not
advertised by the production canary. If a host cannot complete CIMD + S256,
treat that as a release blocker and capture the OAuth trace before considering
a separately reviewed DCR deployment. Do not enable a no-PKCE exception to
bypass connection setup.

The existing confidential compatibility lane is retained for historical
deployments only:

- Set `ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD=client_secret_post`.
- Set `ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE=true`.
- Configure the ChatGPT app with the same predefined client ID, its client
  secret, and token authentication method `client_secret_post`.
- Register the exact ChatGPT callback as an Entra Web redirect.

This fallback is off by default. It preserves the exact client, callback,
resource, scope, and state checks and requires the client secret on token
requests. The public facade-code/refresh-handle store and facade provider
callback are unused in this mode. Rotate the client secret if it may have been
exposed, and prefer the CIMD public lane with mandatory outer and provider S256
PKCE.

The server verifies RS256 signature, issuer, audience, expiry/not-before, and `access_as_user`; it then exchanges the assertion for `https://graph.microsoft.com/.default`. Missing or invalid auth returns both the MCP OAuth challenge metadata and the advertised per-tool `oauth2` scheme. Device-code start/poll/logout tools are disabled inside delegated OAuth requests so a Work user cannot mutate the NAS's legacy shared credential.

For Codex/ChatGPT Work packaging, keep the personal marketplace source at `~/plugins/onedrive` synchronized with this release, run `codex plugin add onedrive@personal`, restart the desktop app, and start a new Work chat. The marketplace entry only installs the package and controls when authentication is prompted; OAuth itself is defined by the mapped ChatGPT app and the Entra registrations above.

### Synology NAS

The DS923+ deployment under `deploy/synology/` runs the MCP server and `tunnel-client` in one Container Manager project. The service uses persistent storage for cache, audit records, backups, and the encrypted Microsoft refresh token; the tunnel runtime key and vault encryption key remain separate owner-only files. No router port forwarding or public MCP listener is required. Work OAuth additionally needs a public HTTPS compatibility origin reverse-proxied to NAS loopback; see `deploy/synology/README.md`.

## Configuration

The setup script writes non-secret config to:

```text
~/.codex/onedrive-plugin/config.json
```

Supported environment variables:

```bash
export ONEDRIVE_CLIENT_ID="your-public-client-app-id"
export ONEDRIVE_TENANT="common"
export ONEDRIVE_SCOPES="offline_access User.Read Files.ReadWrite"
export ONEDRIVE_KEYCHAIN_SERVICE="Codex OneDrive"
export ONEDRIVE_TOKEN_STORE="keychain" # encrypted-file on Linux/NAS
export ONEDRIVE_TOKEN_FILE="$HOME/.codex/onedrive-plugin/auth/tokens.enc"
export ONEDRIVE_TOKEN_ENCRYPTION_KEY_FILE="/run/onedrive-runtime/auth-vault.key"
export ONEDRIVE_STORAGE_ROOT="$HOME/.codex/onedrive-plugin"
export ONEDRIVE_CACHE_ROOT="$HOME/.codex/onedrive-plugin/cache"
export ONEDRIVE_CACHE_TTL_SECONDS="900"
export ONEDRIVE_MAX_SCAN_DEPTH="25"
export ONEDRIVE_MAX_INDEXED_FILE_SIZE="524288"
export ONEDRIVE_INDEX_EXTENSIONS=".txt,.md,.csv,.json,.jsonl,.xml,.yaml,.yml,.html,.css,.js,.mjs,.ts,.tsx,.py,.sql,.log"
export ONEDRIVE_CONCURRENCY_LIMIT="2"
export ONEDRIVE_DELTA_SYNC_ENABLED="true"
export ONEDRIVE_CONTENT_INDEX_ENABLED="true"
export ONEDRIVE_INDEX_OFFICE_EXPORT="false"
export ONEDRIVE_CHATGPT_CACHE_WARM_ENABLED="true"
export ONEDRIVE_CHATGPT_CACHE_WARM_INTERVAL_SECONDS="900"
export ONEDRIVE_CHATGPT_CACHE_WARM_MAX_ITEMS="500"
export ONEDRIVE_CHATGPT_CACHE_WARM_MAX_FOLDERS="100"
export ONEDRIVE_CHATGPT_CACHE_WARM_MAX_DEPTH="8"
export ONEDRIVE_CHATGPT_CACHE_WARM_MAX_PAGES="3"
export ONEDRIVE_PERFORMANCE_LOG="1"

# Optional ChatGPT Work OAuth transport
export ONEDRIVE_MCP_AUTH_MODE="oauth"
export ONEDRIVE_MCP_OAUTH_TENANT="consumers"
export ONEDRIVE_MCP_OAUTH_API_CLIENT_ID="your-mcp-api-app-id"
export ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET_FILE="/run/onedrive-runtime/oauth-api-client.secret"
export ONEDRIVE_MCP_OAUTH_API_RESOURCE="api://your-mcp-api-app-id"
export ONEDRIVE_MCP_OAUTH_API_SCOPE="api://your-mcp-api-app-id/access_as_user"
export ONEDRIVE_MCP_OAUTH_SCOPE_CLAIM="access_as_user"
export ONEDRIVE_MCP_OAUTH_AUDIENCE="your-mcp-api-app-id"
export ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS="your-chatgpt-entra-client-id"
export ONEDRIVE_MCP_OAUTH_AUTHORITY="https://login.microsoftonline.com/consumers/v2.0"
export ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER="https://your-public-oauth-compat-origin.example"
export ONEDRIVE_MCP_OAUTH_OBO_TOKEN_ENDPOINT="https://login.microsoftonline.com/your-validated-tenant-id/oauth2/v2.0/token"
export ONEDRIVE_MCP_OAUTH_GRAPH_SCOPES="https://graph.microsoft.com/.default"
export ONEDRIVE_MCP_PROTECTED_RESOURCE="https://exact-public-tunnel-mcp-endpoint.example"
export ONEDRIVE_MCP_RESOURCE_METADATA_URL="https://exact-public-tunnel-protected-resource-metadata.example"

# Public OAuth compatibility service (map its port to NAS loopback only)
export ONEDRIVE_OAUTH_COMPAT_ENABLED="true"
export ONEDRIVE_OAUTH_COMPAT_HOST="0.0.0.0"
export ONEDRIVE_OAUTH_COMPAT_PORT="3010"
export ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER="https://your-public-oauth-compat-origin.example"
export ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE="$ONEDRIVE_MCP_PROTECTED_RESOURCE"
export ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES="https://tunnel-service.gateway.unified-0.internal.api.openai.org/v1/mcp/tunnel_your-id"
export ONEDRIVE_MCP_PROTECTED_RESOURCE_ALIASES="https://tunnel-service.gateway.unified-0.internal.api.openai.org/v1/mcp/tunnel_your-id"
export ONEDRIVE_OAUTH_COMPAT_CLIENT_ID="your-chatgpt-entra-client-id"
export ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE="/run/onedrive-runtime/oauth-chatgpt-client.secret"
export ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI="https://chatgpt.com/connector/oauth/your-callback-id"
export ONEDRIVE_OAUTH_COMPAT_SCOPES="api://your-mcp-api-app-id/access_as_user offline_access"
export ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD="none"
export ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE="false"
export ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD="true"
export ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR="false"
export ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE="false"
export ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE="/data/auth/oauth-compat-refresh-handles.json"
export ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE="/run/onedrive-runtime/auth-vault.key"
export ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI="https://your-public-oauth-compat-origin.example/callback"
```

`ONEDRIVE_TENANT` can be `common`, `consumers`, `organizations`, or a tenant ID. Use `common` for a plugin that may access either personal Microsoft accounts or work/school accounts.
If Microsoft reports that the app is Microsoft-account-only and requires `/consumers`, the plugin retries device-code and refresh-token auth on `consumers` automatically when the configured tenant is `common`.

`ONEDRIVE_TOKEN_STORE` defaults to `keychain` on macOS and `encrypted-file` on other platforms. The encrypted-file store requires `ONEDRIVE_TOKEN_ENCRYPTION_KEY_FILE` or `ONEDRIVE_TOKEN_ENCRYPTION_KEY`; the key must decode to exactly 32 bytes, and key files with group/other permissions or symlinks are rejected. Prefer the key-file option so the encryption key is not inherited broadly through process environments.

`ONEDRIVE_MCP_AUTH_MODE` defaults to `noauth`. Set it to `oauth` only after the Entra API registration, OBO secret, API resource/scope/audience, allowed ChatGPT client ID, public protected-resource URLs, public compatibility origin, private upstream Entra client registration, facade callback, and ChatGPT callback are configured. In OAuth mode both processes fail closed at startup if required settings, key files, or secret files are missing, identifiers retain deployment placeholders, or a public URL is not HTTPS. The API secret is used only for OBO. The separate compatibility secret is used only for the facade's private upstream Entra code/refresh exchange; outer ChatGPT authentication is `none` plus S256 PKCE. Neither secret is returned, logged, or written to audit records.

Every ChatGPT tool result includes safe `_meta["onedrive/performance"]` phase totals for server, auth, Graph, verification, and cache work; the same redacted values are written to structured logs. Tunnel and ChatGPT model-selection time occur outside the MCP server, so compare the host-observed duration with `serverMs` to estimate that external portion. OAuth cache warming runs only in the background, is keyed to the authenticated drive subject, uses delta refresh when possible, and is bounded by the `ONEDRIVE_CHATGPT_CACHE_WARM_*` settings.

You can also add friendly path aliases to the config file:

```json
{
  "pathPresets": {
    "job-tracker": "Documents/Job Tracker",
    "screenshots": "Pictures/Screenshots"
  }
}
```

Tools that accept `path` can also accept `preset` plus `relativePath`; upload/write tools accept `remotePreset` plus `remoteRelativePath`.

Optional performance settings can also be placed in the config file:

```json
{
  "storageRoot": "/absolute/path/to/onedrive-plugin",
  "cacheRoot": "/absolute/path/to/onedrive-plugin/cache",
  "settings": {
    "cacheTtlSeconds": 900,
    "maxScanDepth": 25,
    "concurrencyLimit": 2,
    "deltaSyncEnabled": true,
    "contentIndexEnabled": true
  },
  "indexing": {
    "maxFileSize": 524288,
    "supportedExtensions": [".txt", ".md", ".csv", ".json", ".jsonl", ".xml", ".yaml", ".yml"],
    "includeOfficeTextExport": false
  }
}
```

Use absolute paths in `storageRoot` and `cacheRoot` if you override them. Configuration precedence is environment variables, then `~/.codex/onedrive-plugin/config.json`, then server defaults. The checked-in `.mcp.json` does not override tenant or scopes.

## Tools

- `onedrive_config`
- `onedrive_doctor`
- `onedrive_auth_device_start`
- `onedrive_auth_device_poll`
- `onedrive_logout`
- `onedrive_me`
- `onedrive_drive`
- `onedrive_presets`
- `onedrive_list`
- `onedrive_list_all`
- `onedrive_scan`
- `onedrive_search`
- `onedrive_search_all`
- `onedrive_find`
- `onedrive_find_all`
- `onedrive_delta`
- `onedrive_sync_status`
- `onedrive_cache_refresh`
- `onedrive_cache_clear`
- `onedrive_content_index_refresh`
- `onedrive_content_search`
- `onedrive_content_index_clear`
- `onedrive_office_capabilities`
- `onedrive_office_validate`
- `onedrive_office_index_refresh`
- `onedrive_office_search`
- `onedrive_word_get_document`
- `onedrive_excel_get_workbook`
- `onedrive_powerpoint_get_presentation`
- `onedrive_word_batch_update`
- `onedrive_excel_batch_update`
- `onedrive_powerpoint_batch_update`
- `onedrive_office_batch_transform`
- `onedrive_office_backups`
- `onedrive_office_compare_backup`
- `onedrive_office_restore_backup`
- `onedrive_get_info`
- `onedrive_read_text`
- `onedrive_preview`
- `onedrive_download`
- `onedrive_download_excel`
- `onedrive_download_word`
- `onedrive_download_powerpoint`
- `onedrive_export_pdf`
- `onedrive_export_text`
- `onedrive_upload`
- `onedrive_write_text`
- `onedrive_patch_text`
- `onedrive_versions`
- `onedrive_compare_version`
- `onedrive_restore_version`
- `onedrive_workspace_list`
- `onedrive_workspace_create`
- `onedrive_workspace_status`
- `onedrive_workspace_promote`
- `onedrive_workspace_abandon`
- `onedrive_watch_start`
- `onedrive_watch_status`
- `onedrive_watch_stop`
- `onedrive_create_folder`
- `onedrive_rename`
- `onedrive_move`
- `onedrive_copy`
- `onedrive_create_sharing_link`
- `onedrive_invite_permission`
- `onedrive_revoke_permission`
- `onedrive_batch_revoke_permissions`
- `onedrive_permissions`
- `onedrive_batch_get_info`
- `onedrive_batch_permissions`
- `onedrive_batch_download`
- `onedrive_batch_delete`
- `onedrive_batch_move`
- `onedrive_update_file`
- `onedrive_recent`
- `onedrive_large_files`
- `onedrive_duplicates`
- `onedrive_shared_by_me`
- `onedrive_public_links`
- `onedrive_restore_deleted`
- `onedrive_audit_recent`
- `onedrive_audit_export`
- `onedrive_audit_clear`
- `onedrive_delete`

The focused ChatGPT profile advertises 19 tools. Route exact filenames, known root-relative file paths, or observed HTTPS OneDrive/SharePoint file links through `onedrive_open_files`, broader or multi-part reads through `onedrive_read_actions`, and selected readable results through `fetch`. For `onedrive_open_files`, pass exactly one of `names` or `urls`, with at most five values, and present each returned `displayLink` so only the resolved filename is visible as hyperlink text rather than the bare URL. Use `onedrive_office_inspect` for bounded Word/Excel/PowerPoint structure, `onedrive_office_capabilities` for one exact operation schema and valid example, `onedrive_office_review` for Word comments or Excel notes at exact evidence coordinates, and `onedrive_office_batch_transform` for other structured edits. Use `onedrive_download_file` when the user needs an original or PDF as an MCP resource, and `onedrive_render_preview` when pages or slides need visual QA as images. Rename/move/copy/share/revoke/version-restore changes go through `onedrive_preview_actions`, then `onedrive_commit_actions` after approval. Direct create, upload, export, text-write/patch, invite, recycle-bin delete, and recycle-bin restore requests use their named focused tools.

## Native Office editing

The plugin can inspect and edit modern Open XML files without requiring Word, Excel, or PowerPoint to be open. Reads expose structured document content and package-safety metadata. Mutation previews include semantic operation counts and affected objects, are bound to file identity and eTag/cTag, backed up locally by default, uploaded with `If-Match`, validated after commit, and recorded in the mutation audit.

Managed Office backups have opaque IDs and manifests containing the original stable item ID and version metadata. Trusted local full-profile clients use `onedrive_office_backups` to list them and `onedrive_office_compare_backup` for a bounded semantic comparison with the current remote content. Focused hosted clients list backups through `onedrive_read_actions` operation `officeBackups`, then remove exact selected IDs only through preview/commit operation `deleteOfficeBackups`. `onedrive_office_restore_backup` defaults to dry-run and requires the preview token, original `expectedId`, and current `expectedETag`; it restores by item ID, creates a rollback backup, audits the mutation, and verifies the restored fingerprint.

- Word exposes durable paragraph, table, content-control, and comment evidence anchors and 22 headless operations, including comment creation/deletion, image insertion/replacement, bookmarks, content controls, table row/column changes, headers/footers, and section properties. `deleteComment` removes the comment plus its range/reference markers. Documents containing tracked changes are refused instead of silently changing review semantics.
- Excel exposes worksheet, range, defined-name, and table/row-key anchors and 33 headless operations, including worksheet/table lifecycle, merge/unmerge, sort/filter, hyperlinks, notes, images, chart formatting, passwordless sheet protection, and pivot refresh-on-open. Business `.xlsx` Graph writes use scoped persistent sessions; personal workbooks and unsupported Graph operations use Open XML automatically.
- Open XML validation now rejects undeclared prefixes referenced by `mc:Ignorable`, and serialization preserves original namespace declarations even when a prefix appears only in compatibility metadata. `addTableRow` copies the nearest row's cell styles and row attributes/height and extends applicable conditional-format/data-validation ranges. Its canonical `values` payload is a two-dimensional rows array; a one-dimensional single-row shorthand is normalized before validation and execution. `deleteTableRow` compacts one table data row, shifts native hyperlinks, translates ordinary relative A1 formulas, safely shrinks bounded single-column shared-formula groups without changing their formulas or cached values, shrinks applicable ranges, and reports each preservation decision in its preview diff.
- Untrusted Open XML packages are bounded before parsing: at most 10,000 entries, 256 MiB total expansion, 64 MiB per ZIP member, and 32 MiB per XML/relationship part. Fingerprints stream package members, and Linux helper processes apply CPU, output-file, and 512 MiB address-space ceilings in addition to the container limit.
- PowerPoint exposes persistent slide and shape anchors and 25 headless operations, including slides, images/cropping, tables and row/column changes, alternative text, z-order, grouping/ungrouping, and layout application.
- Positional selectors remain valid. An anchor defaults to `rebasePolicy:"unique"`; moved targets re-resolve only when exactly one match exists. Missing, duplicate, or selector/anchor disagreement returns a structured conflict. Every live commit remains bound to the eTag used by its preview.
- Encrypted and legacy binary files are refused. Macro-enabled edits require `allowMacros: true`; signed-package edits are always refused because any edit invalidates the signature.

In the focused profile, call `onedrive_office_inspect` first and preserve its stable item ID and eTag, then request only the intended operation from `onedrive_office_capabilities`. Call the batch-transform or review tool as a dry run before a live commit. A live commit requires `dryRun: false`, `confirmed: true`, `expectedName` or `expectedId`, and the exact returned `previewToken`; evidence-bound review actions additionally require the exact paragraph/comment or worksheet/address coordinate returned by inspection.

For faster research inside a known Office file, use the focused inspector's bounded Word paragraph/table selectors, Excel `sheetNames` and A1 `address`, or PowerPoint slide/shape selectors. The full-profile structured reads also accept `searchText`; Excel can restrict reads to `sheetNames` and a bounded A1 `address`.

For cross-drive research, `onedrive_office_index_refresh` stores structured paragraph, cell, formula, table, content-control, comment, shape, and notes segments with semantic anchors. It reuses unchanged eTag/cTag entries and prefers the existing OneDrive delta cursor before scanning. `onedrive_office_search` searches that private local index without Graph calls. `onedrive_office_batch_transform` preflights every requested file before the first write and returns recovery backup IDs if a later item fails.

## Remote editing workflows

- `onedrive_versions`, `onedrive_compare_version`, and `onedrive_restore_version` expose bounded Graph version history, semantic/text/binary comparisons, and native restore. Restore never falls back to replacing content and requires current identity, eTag, confirmation, and its preview token.
- The focused `onedrive_read_actions` composite exposes `recent`, `versions`, and `compareVersion`; `restoreVersion` is available only through the separate preview/commit composites and remains bound to the selected version, current eTag, item identity, and preview proof.
- Focused `drives`, `enterpriseSearch`, and `libraryList` reads discover accessible work/school libraries and keep an explicit `driveId` attached to every subsequent library lookup. Cross-library mutation is not implied. The deprecated `/me/drive/sharedWithMe` endpoint is not used.
- `onedrive_download_file` materializes a bounded original or Graph-converted PDF behind a short-lived, scope-bound MCP resource URI. `onedrive_render_preview` converts supported Office/PDF content to bounded page or slide images for visual QA. Neither focused tool exposes a NAS/local filesystem path.
- `onedrive_patch_text` applies bounded unified diffs, RFC 6902 JSON Patch, restricted safe-YAML path operations, or RFC 4180 CSV row-key operations. It preserves supported BOM/encoding, newline style, and trailing-newline state while refusing binary and oversized inputs.
- `onedrive_workspace_*` manages owner-only drafts under `Codex Editing Drafts`. Workspaces record the original stable item ID and base eTag/version, surface source/draft drift, block promotion after source drift, preserve the original item identity/version history on success, and retain failed/conflicted drafts for recovery.
- `onedrive_watch_*` manages auth-context/drive-scoped delta watches with 15–300 second polling, one-hour default expiry, eight-hour maximum, throttling backoff, and a 500-event ring buffer. Events invalidate affected previews and mark source workspaces stale.

## Safety

- Remote mutations that move, rename, copy, expose, invite, restore, delete, or revoke access use a preview-first pattern.
- Rename, move, copy, sharing-link creation, named-recipient invitation, permission revoke, restore, and delete default to dry-run where the operation has a dry-run mode.
- Live rename, move, copy, sharing-link creation, named-recipient invitation, permission revoke, restore, and delete require `dryRun: false`, `confirmed: true`, and stable expected identity (`expectedName` or `expectedId`; restore requires `expectedId`).
- Replacing an existing file with `onedrive_upload` or `onedrive_write_text` is also preview-token gated: review the returned existing item, then repeat with `dryRun: false`, `confirmed: true`, matching expected identity, and the exact preview token.
- Live sharing-link creation, named-recipient invitation, permission revoke, batch permission revoke, restore, delete, and batch delete also require the `previewToken` returned by the immediately preceding dry-run preview for the same resolved operation.
- Batch delete, batch move, and batch permission revoke preflight every item before any mutation and refuse partial execution when a preflight check fails. Live batch responses include a warning that successful earlier items may already be changed if a later item fails.
- Sharing-link creation supports Microsoft Graph link type/scope plus optional password and expiration, and can include a before/after permission diff so the caller can see what changed.
- `onedrive_invite_permission` grants named users or groups access through Microsoft Graph `driveItem: invite`. It defaults to a silent direct grant (`sendInvitation: false`, `requireSignIn: true`); email invitations are opt-in with `sendInvitation: true` and optional `message`.
- Permission revoke uses Microsoft Graph `DELETE /me/drive/items/{item-id}/permissions/{permission-id}` and includes before permissions by default; live revoke includes after permissions and a permission diff.
- Rename, move, copy, share, and delete refuse to operate on the OneDrive root.
- A relative target or destination field is refused unless its matching preset field is present; it never silently resolves relative to the OneDrive root.
- Tool arguments are validated before handlers run, including required fields, unknown properties, enum values, numeric bounds, array bounds, and target `anyOf` rules.
- Text reads are bounded to 5 MB by default.
- Text reads use MIME/extension checks and refuse likely binary files unless `force: true` is set.
- Full-profile downloads go to an opaque authentication-subject-and-drive subdirectory below `~/.codex/onedrive-plugin/downloads` unless `localPath` is provided. Concurrent default downloads, exports, update checkouts, backups, and audit exports reserve unique local paths instead of racing on the same filename.
- Downloads and uploads refuse local OneDrive sync-folder paths by default. Use `allowLocalOneDriveSyncPath: true` only for an explicit local sync-folder workflow.
- Uploads use simple upload for smaller files and upload sessions for large files, or when `uploadMode: "session"` is requested.
- List, search, find, scan, and delta tools return compact item summaries by default; pass `format: "full"` for richer metadata.
- Normal list, search, scan, delta, and metadata calls opportunistically maintain a local metadata cache in an opaque scope directory below `~/.codex/onedrive-plugin/cache`.
- Plugin-managed cache, index, audit, download, Office-backup/editing, upload-staging, workspace, watch, export, and update-workflow state is physically separated by an opaque SHA-256 scope derived from the delegated OAuth subject/authentication context plus active drive ID. In-memory caches, previews, fetch continuations, and watch timers use the same scope boundary. Managed directories are restricted to the current user (`0700`), and persisted files are restricted to the current user (`0600`).
- `onedrive_sync_status` reports cache age, item count, delta cursor availability, resumable delta next-link availability, unresolved path count, and plugin storage locations.
- Metadata cache v4 and content/Office indexes v3 are bound to both an opaque authentication context and the active drive ID. Delta cursors, preview tokens, fetch snapshots, Office backups, workspaces, watches, upload staging, downloads, and update manifests use the same physical and in-memory scope; legacy unscoped local state is invalidated or retained but refused until safely recreated. Local-only search fails closed on a scope mismatch. Cache, index, and audit files use atomic/locked updates with reload-on-write freshness checks.
- `onedrive_cache_refresh` rebuilds the cache from a bounded recursive scan and uses delta refreshes when a previous cursor exists for the same root. Cache refresh batches metadata-cache writes during scans, persists only delta-origin `nextLink` cursors for continuation, reconciles pathless delta records through cached parent IDs, and returns progress milestones. Ordinary list/search pagination cannot seed delta state. `onedrive_cache_clear` clears the cache.
- `onedrive_content_index_refresh` is the explicit content-reading step. It indexes supported cached text and structured Office content into `content-index.json`, stores normalized text/tokens plus semantic Office anchors, reuses entries when ETag/cTag/mtime/size are unchanged, and applies file-size, segment, concurrency, and per-item failure limits. Explicit metadata deletes and changed fingerprints evict stale entries; moves and renames update indexed metadata; unchanged explicit cTags preserve content entries across metadata-only renames, while changed or omitted content tags with changed ETags invalidate conservatively. Bounded partial scans do not globally prune unseen entries.
- `onedrive_content_search` searches only the local content index and returns lightweight metadata plus snippets. It does not call Microsoft Graph or read file bodies.
- The focused ChatGPT profile advertises `onedrive_read_actions` for folder listings, descriptive search, item metadata, permission inspection, recent items, version history/comparison, and explicit enterprise drive/library discovery. `onedrive_open_files` handles one to five exact filenames, known root-relative file paths, or observed HTTPS OneDrive/SharePoint file links in one bounded read-only call. Its model-visible content is filename-only Markdown hyperlinks, while canonical URLs and stable IDs remain in structured content for follow-up use. Known paths bypass search; validated links resolve through Microsoft Graph share addressing to the exact source drive; multiple filename index misses share one capped recursive traversal instead of rescanning the same nested folders for every file. `fetch` accepts a result ID and returns at most 192 KiB of readable text. It directly extracts structured `.docx`/`.xlsx`/`.pptx` content; reads CSV/TSV, JSON, XML, Markdown, HTML, source code, and other text formats; and supports bounded local extraction for PDF, RTF, OpenDocument, EPUB, legacy `.doc`/`.xls`/`.ppt`, and common images when the deployment extractor is available. The NAS image includes the required PDF, OCR, and legacy Office extractors. All ChatGPT tool payloads are bounded to 1 MiB. The full 84-tool profile remains available through trusted local stdio only; hosted HTTP rejects it because it includes intentional local-filesystem maintenance capabilities.
- `onedrive_preview_actions` batches up to ten rename, move, copy, sharing-link, permission-revoke, or native version-restore previews as one read-only ChatGPT call. It never performs a mutation, returns the operation-bound tokens needed by the separate live tools, and emits only permission counts/link counts/roles for sharing previews. Names, emails, permission objects, and sharing URLs are omitted from this preview result. Live sharing-link creation is separately marked as open-world and still requires explicit approval, exact expected identity, and the preview token.
- `onedrive_find` and `onedrive_find_all` can merge local content-index hits into ranking, but they never fetch or parse full content themselves. Build or refresh the index first when content search is needed.
- `onedrive_find` is the preferred file lookup helper. It uses the local metadata cache when available, confirms exact strong cache hits with live metadata, runs the canonical Graph query first, and expands additional terms in bounded concurrent waves only while confidence remains low. Canonical Graph results can represent filename, metadata, or file-content matches; unrelated expansion-only results remain gated. Results expose the planned, executed, and skipped search terms. `graphSearchCalls` reports actual Graph search pages fetched, not just term count. Tune expansion with `searchConcurrency` and fallback scans with `scanConcurrency`. Fallback scans prune duplicate and nested folder hints regardless of input order. Cache-only hits must still have query relevance and are not treated as authoritative when live evidence cannot confirm them. Pass `useCache: false` for a fully live lookup with no metadata-cache reads or writes.
- `onedrive_find_all` is the broader locator for “look everywhere” requests. It searches every planned term instead of stopping after the first confident canonical result, searches common folders first, and uses larger bounded scan caps, with cache acceleration and the same duplicate-hint pruning when available.
- `onedrive_preview` returns bounded text previews for text files and Graph-supported document text exports without reading unbounded remote content into memory.
- `onedrive_update_file` provides a checkout/commit edit workflow with a local manifest, eTag/cTag/size/mtime conflict checks, optional backup, and post-commit verification. Checkout refuses to overwrite an existing manifest unless `overwriteManifest: true` is provided.
- `onedrive_batch_get_info` and `onedrive_batch_permissions` use Microsoft Graph batching for up to 20 items. Batch download/delete/move tools provide one result per item with dry-run support where destructive.
- `onedrive_rename`, `onedrive_move`, and `onedrive_copy` support `dryRun: true` previews. Live `onedrive_batch_move` requires `dryRun: false`, `confirmed: true`, and `expectedName` or `expectedId` for every item.
- `onedrive_recent`, `onedrive_large_files`, `onedrive_duplicates`, `onedrive_shared_by_me`, and `onedrive_public_links` provide cleanup and sharing-audit workflows.
- `onedrive_list_all` follows pagination within one folder. Use `onedrive_scan` when you need recursive traversal across subfolders or the whole OneDrive. Direct scans use bounded folder concurrency; set `scanConcurrency` from 1–4 when latency or Graph throttling requires an explicit tradeoff.
- `onedrive_doctor` checks config, auth, profile, drive metadata, every configured preset target, and optional root listing in one call. Missing preset folders produce a warning with the exact aliases that need a `pathPresets` override.
- Tool failures retain a short text message for compatibility and also return machine-readable `structuredContent.error` metadata with a stable code such as `not_found`, `permission_denied`, `conflict`, `rate_limited`, or `service_unavailable`.
- `onedrive_export_pdf` and `onedrive_export_text` ask Microsoft Graph to convert supported Office files before saving locally. Microsoft Graph may reject conversions for unsupported file types.
- `onedrive_permissions` audits current sharing/permission grants before changing access.
- `onedrive_delta` can return deleted item changes. Set `maxPages` from 1 to 100 to cap Graph pages in one call while retaining the advanced `nextLink` or terminal `deltaLink`. Microsoft Graph does not expose a normal OneDrive recycle-bin listing endpoint through the driveItem file APIs.
- `onedrive_get_info` supports `includeDeletedItems: true` when targeting an item ID; Microsoft documents this as OneDrive Personal-only.
- `onedrive_restore_deleted` defaults to dry-run and requires a deleted item ID. Live restore may require `Files.ReadWrite.All` for personal OneDrive.
- Live remote mutations are recorded in `mutations.jsonl` inside the current opaque subject-and-drive subdirectory below `~/.codex/onedrive-plugin/audit`. Audit entries include safe item summaries, before/after summaries when available, permission diffs when relevant, Graph request IDs when available, and safe error details for failed live mutations. They do not log tokens, authorization headers, file contents, raw request bodies, sharing-link web URLs, passwords, invite messages, or recipient identifiers.
- A successful remote mutation remains a success if local cache/audit bookkeeping or a best-effort post-mutation verification later fails. The response reports `localWarnings` and, where applicable, `verificationIncomplete`; do not repeat the mutation merely to repair that follow-up state.
- `onedrive_audit_recent` reads recent audit entries, `onedrive_audit_export` exports the JSONL log to a local file, and `onedrive_audit_clear` requires `confirmed: true`.
- Graph requests retry transient `429`, `500`, `502`, `503`, and `504` responses with `Retry-After` support. Read-only requests also retry transient transport failures, and Microsoft Graph batch helpers retry only the transient individual subrequests while preserving result order.

## Performance Architecture

The plugin separates cheap metadata discovery from expensive content reads:

- Metadata cache: stores IDs, drive/item metadata, paths, web URLs, MIME/type hints, size, timestamps, ETag/cTag, and file/folder status. This avoids repeated full recursive scans when cached metadata is fresh enough for the workflow.
- Delta sync: `onedrive_cache_refresh` prefers stored Microsoft Graph delta cursors when the requested root matches the cached root. If Graph returns a delta `nextLink` before the final `deltaLink`, the plugin stores that incomplete cursor and resumes it on the next refresh. It resolves pathless delta records from cached parent IDs where possible, reports unresolved path counts for parents it cannot hydrate, repaths descendants after folder moves/renames, removes deleted descendants, and rejects legacy non-delta cursors. It falls back to a bounded scan when no cursor exists, the target changed, delta is disabled, or Graph rejects the cursor.
- Search ranking: `onedrive_find` combines exact path/filename evidence, canonical live Graph content/metadata matches, confirmed metadata-cache matches, folder hints, file-type hints, recency-ish modified metadata, and optional content-index hits. Adaptive term execution stops expansion after a confident live result for normal `find`; `find_all` executes all planned terms for exhaustive locator requests. Results include reasons and request-plan counters so ranking and latency are debuggable.
- Content index: indexing is opt-in and explicit. It supports bounded text-like files by default, stores normalized text/tokens for faster repeated queries, keeps only bounded top matches during local content search, and can optionally try Graph `format=text` export for Office-like files. Large files, unsupported binaries, failed exports, and files over the cap are skipped or reported without aborting the whole refresh. Cache reconciliation removes demonstrably stale entries without treating a bounded partial scan as a complete-drive deletion signal.
- Graph optimization: once an item is discovered, the plugin prefers item IDs for follow-up reads/mutations, follows pagination with cycle/page caps, uses `$batch` where useful, retries transient throttling/service errors (including transient batch subresponses), and honors `Retry-After`.

The largest performance risk is any broad recursive scan or content-index refresh over a large OneDrive tree. Keep `maxItems`, `maxFolders`, `maxDepth`, `maxFiles`, and `maxBytesPerFile` bounded, and warm the metadata cache before broad repeated searches.

## Troubleshooting and Benchmarks

Run a health check:

```bash
onedrive_doctor({ "checkRootList": true })
```

Inspect cache/index state:

```bash
onedrive_sync_status({ "includeSamples": true })
```

Refresh metadata using scan/delta:

```bash
onedrive_cache_refresh({ "mode": "auto", "maxItems": 10000, "maxFolders": 2000 })
```

Build the optional content index from cached metadata:

```bash
onedrive_content_index_refresh({ "maxFiles": 100, "maxBytesPerFile": 524288 })
```

For a simple before/after benchmark, compare:

1. Cold search: clear cache, run `onedrive_find` with `useCache: false`.
2. Warm metadata search: run `onedrive_cache_refresh`, then repeat `onedrive_find`.
3. Content-indexed search: run `onedrive_content_index_refresh`, then run `onedrive_content_search` and `onedrive_find` for a phrase inside a file.
4. Selected file read: use `onedrive_preview` or `onedrive_read_text` only after selecting a specific result.

The bundled benchmark script runs those steps through the MCP server with bounded caps:

```bash
scripts/benchmark.mjs --query="project plan" --maxItems=1500 --maxFolders=250 --maxFiles=50 --searchConcurrency=2
```

Validate the ChatGPT-specific metadata budget separately from Microsoft Graph latency:

```bash
node scripts/tool-profile-test.mjs
node scripts/chatgpt-golden-test.mjs
```

The guards preserve the full 84-tool contract for Codex, limit ChatGPT to the reviewed 19-tool allowlist, and reject direct calls to hidden full-profile handlers. They verify exact composite-operation enums, targeted Office capability schemas/examples, read-only/destructive annotations, the 38 KiB OAuth `tools/list` budget, bounded results, and every focused skill's declared tool references. The golden-prompt gate covers every focused tool and checks the metadata cues that distinguish commonly confused actions. The ChatGPT server version includes a deterministic contract hash so ChatGPT invalidates stale tool metadata whenever the advertised surface changes.

The hosted MCP server also advertises the `io.modelcontextprotocol/skills` extension. `skills/list`, `skills/get`, and exact `skill://onedrive/...` resource reads expose the five bundled OneDrive workflow skills and their SHA-256-verified supporting files for ChatGPT's **Scan Tools** import. The catalog is static and readable during plugin scanning; private OneDrive resources and all tools remain OAuth-protected. Re-run **Scan Tools** on the existing app after deployment so ChatGPT imports the updated snapshot without creating a duplicate app.

Add `--clear` when you intentionally want to clear local metadata/content caches before the cold run. The script performs read-only Microsoft Graph operations, writes local cache/index files, and emits progress events to stderr while keeping the final summary JSON on stdout.

Expected improvement: exact multi-file reads now require one host tool round instead of a search/fetch round for every file, and multi-action previews require one read-only host round instead of one consent-classified call per action. Repeated focused folder lists reuse a subject-scoped in-memory snapshot for up to 60 seconds and every successful mutation invalidates that subject's snapshots. Descriptive search actions in the combined ChatGPT Work read tool now use ranked discovery instead of returning raw Graph order. Exact filenames, complete filename phrases, and full filename/path token coverage form higher relevance tiers than verified content matches; unverified content-only Graph hits are suppressed. Single-token queries perform a live search rather than stopping on one incomplete cache candidate, while low-confidence cold discovery can use a bounded metadata scan. A high-confidence multi-token match in a fresh metadata cache returns without a Graph search. A high-confidence match in a cache no more than 24 hours old also returns immediately while a bounded delta or query refresh runs in the background. Cold ChatGPT discovery searches up to six evidence-ranked terms in waves of three: the natural-language query, any specific names/numbers, and deterministic aliases for common home-service, medical, tax, insurance, travel, employment, and agreement intents. Corroborating concept hits receive a bounded ranking boost, while generic report/service words alone cannot outrank stronger domain evidence. Fetch validates a stale content-index entry with lightweight metadata and reuses it when the ETag/cTag fingerprint is unchanged. Large documents return a representative first response capped at 32 KiB, with `metadata.nextChunkId` for sequential 64 KiB continuation reads through the same `fetch` tool; continuations reuse a short-lived in-memory extraction when possible. Actual speed still depends on ChatGPT host latency, Microsoft Graph latency, OneDrive size, throttling, and configured caps.

## Safe Example Prompts

- "Find the file named Project Plan, show me its item ID and current permissions, but do not change anything."
- "Dry-run moving this file to Documents/Archive and show the exact item ID I would need to confirm."
- "Dry-run granting person@example.com read access to this item without sending an email."
- "Revoke the anonymous link from this item only after previewing the permission diff and asking me to confirm."
- "Upload this attached Word document to Documents, preview the exact destination first, and do not replace anything."
- "Move this item to the recycle bin only after showing me its exact ID and asking me to confirm."
- "Restore this deleted personal-OneDrive item by ID; preview the destination and ask before restoring it."
- "Permanently delete this item only after an irreversible-action preview, exact identity check, and a second explicit confirmation."
- "Show recent OneDrive mutation audit entries from this plugin."

## Beta Test

Run the mocked Microsoft Graph regression suite first. It does not touch OneDrive, Keychain, or Microsoft services:

```bash
scripts/mock-graph-test.mjs
```

Run the focused ChatGPT Work beta against the connected account after the mocked suite. This is read-only: it starts the exact 19-tool Work profile, measures ranked search quality for representative metadata and exact-filename queries, fetches a known document, writes an evidence report, and makes no remote mutations:

```bash
node scripts/chatgpt-work-beta-test.mjs --run-id=codex-beta-work-YYYYMMDDTHHMMSSZ --report=work/qa-artifacts/codex-beta-work-YYYYMMDDTHHMMSSZ.json
```

Run the benchmark script when comparing cold search, warm cache search, indexed content search, and selected preview timing. The script exits nonzero if any MCP tool step reports an error:

```bash
scripts/benchmark.mjs --query="project plan"
```

Run the prepackage guard before refreshing the plugin cache:

```bash
scripts/prepackage-check.mjs
```

Preview the exact new versioned cache directory, then install only after reviewing that path. The installer runs the source prepackage gate, creates the version directory atomically, refuses any existing target, preserves older cache versions, and requires full byte/mode/type/symlink parity before succeeding:

```bash
node scripts/install-versioned-cache.mjs
node scripts/install-versioned-cache.mjs --confirmed --target="$HOME/.codex/plugins/cache/personal/onedrive/0.6.4+codex.20260816170509"
```

After both live betas, regenerate the two QA reports, preview their exact sync into that new cache, then apply only those evidence files and re-run parity:

```bash
node scripts/install-versioned-cache.mjs --sync-evidence --target="$HOME/.codex/plugins/cache/personal/onedrive/0.6.4+codex.20260816170509"
node scripts/install-versioned-cache.mjs --sync-evidence --confirmed --target="$HOME/.codex/plugins/cache/personal/onedrive/0.6.4+codex.20260816170509"
```

Office compatibility checks are split by purpose:

```bash
python3 scripts/office-openxml-test.py
python3 scripts/office-security-test.py
node scripts/semantic-anchors-test.mjs
node scripts/text-patch-test.mjs
SOFFICE="$(command -v soffice)" python3 scripts/office-real-fixture-test.py
```

Install the pinned fixture dependencies with `python3 -m pip install -r scripts/requirements-office-test.txt`. The security corpus includes malformed ZIP/XML/relationship cases, deterministic mutation fuzzing, every possible two-run PowerPoint split of a target phrase, and a 5,000-run deck. The real-fixture gate generates packages with `python-docx`, `openpyxl`, and `python-pptx`, edits them, reopens them with their native libraries, and requires LibreOffice PDF conversion without repair/corruption diagnostics.

After installing a refreshed build, compare source with the installed cache:

```bash
scripts/prepackage-check.mjs --installed /path/to/installed/onedrive/cache
```

Running the harness without `--live` is read-only. It prints the exact proposed run ID and folder name. Review those values, then run the live CRUD/regression test from the plugin directory or with an absolute path to the installed plugin:

```bash
scripts/beta-test.mjs --live --confirmed --run-id=codex-beta-20260713t150000z --invite-recipient=person@example.com
```

For a durable machine-readable result, add `--report=work/qa-artifacts/<run-id>.json`. When a report path is provided, the harness writes the complete JSON result there and keeps terminal output compact; the ignored `work/` tree prevents local beta evidence from entering the plugin package.

All four live arguments are required. The test creates the exact named temporary OneDrive folder, exercises CRUD and safety behavior, silently grants then revokes read access for the explicit recipient, creates then revokes an anonymous test link, deletes only that test folder during cleanup, and removes isolated local work on success. Results are recorded as `pass`, `fail`, or `blocked`; a resource limitation is never reported as a pass. Pass `--keep-work` to keep local artifacts for debugging.
The invite recipient must differ from the signed-in OneDrive owner. The harness rejects a self-recipient before creating the remote test folder because Microsoft returns the existing non-revocable owner grant rather than a new isolated permission; an explicit user-controlled email alias is acceptable when the provider treats it as a distinct recipient.
If Microsoft Graph rejects a distinct recipient with `sharingFailed`, the harness immediately audits permissions. It continues with that named-recipient round trip marked blocked only when the audit proves no grant was created; any partial or ambiguous permission state remains a hard failure.
The live harness uses one Python runtime for both Office fixtures and the plugin child process: `ONEDRIVE_OFFICE_TEST_PYTHON`, then `ONEDRIVE_OFFICE_PYTHON`, then `python3` from `PATH`. Set `ONEDRIVE_OFFICE_TEST_PYTHON=/absolute/path/to/python` when the pinned fixture dependencies are installed in another Python environment. Fixture setup failures are reported with the exact remediation command and remove partial local beta work unless `--keep-work` was requested; they occur before any remote mutation.
Unknown, duplicate, and positional CLI options are rejected before the beta harness starts, so a misspelled safety or mode flag cannot silently fall through to a different test mode.
The harness emits one compact progress event to stderr after every completed check. Live child Graph requests use a 10-second timeout by default and the harness does not multiply the server's built-in read retries; use `--fetch-timeout-ms=1000..60000` or `--read-retry-attempts=1..5` only when intentionally tuning a degraded connection.

Find old beta-test folders without deleting them. Cleanup discovery uses bounded Graph search followed by item verification, so it does not recursively scan the entire drive. Candidates with missing or invalid timestamps are skipped, and invalid/overflowing cleanup limits are rejected before any delete:

```bash
scripts/beta-test.mjs --cleanup-stale --stale-days=1
```

Delete the stale candidates only after reviewing the dry-run output:

```bash
scripts/beta-test.mjs --cleanup-stale --stale-days=1 --live --confirmed --run-id=codex-beta-cleanup-20260713
```

Run read-only tenant health checks across personal/work-school tenant endpoints:

```bash
scripts/beta-test.mjs --tenant-matrix=common,consumers,organizations
```

Use `--tenant-matrix-live --live --confirmed --run-id=<exact-id> --invite-recipient=<email>` only when you intentionally want to run the full live beta once per tenant entry.

## Plugin Gallery

The plugin manifest includes a file-manager flow screenshot at `assets/screenshot-file-manager.png` so the Codex plugin page shows the OneDrive search, read, upload, and safety workflow rather than only an icon.

![OneDrive plugin file manager flow](assets/screenshot-file-manager.png)

## CI

The GitHub Actions workflow in `.github/workflows/ci.yml` runs syntax checks, pinned Office fixture/security checks, semantic-anchor and structured-patch tests, task-pane and loopback-broker security tests, the Microsoft Graph regression suite, and the prepackage guard on Node.js 24 and 26 for every push and pull request.

## Microsoft References

- OneDrive files in Microsoft Graph: https://learn.microsoft.com/en-us/graph/api/resources/onedrive
- List folder children: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children
- Large file upload sessions: https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
- Move driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-move
- Copy driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-copy
- Create sharing link: https://learn.microsoft.com/en-us/graph/api/driveitem-createlink
- Invite recipients: https://learn.microsoft.com/en-us/graph/api/driveitem-invite
- Drive recipient resource: https://learn.microsoft.com/en-us/graph/api/resources/driverecipient
- Delta sync: https://learn.microsoft.com/en-us/graph/api/driveitem-delta
- Search drive items: https://learn.microsoft.com/en-us/graph/api/driveitem-search
- JSON batching: https://learn.microsoft.com/en-us/graph/json-batching
- driveItem resource fields: https://learn.microsoft.com/en-us/graph/api/resources/driveitem
- List permissions: https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions
- Get driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-get
- Restore deleted item: https://learn.microsoft.com/en-us/graph/api/driveitem-restore
- Microsoft Graph throttling guidance: https://learn.microsoft.com/en-us/graph/throttling
- Delegated Microsoft Graph auth: https://learn.microsoft.com/en-us/graph/auth-v2-user
- Microsoft identity on-behalf-of flow: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow
- Expose an Entra-protected web API: https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-configure-app-expose-web-apis
- OpenAI Apps SDK authentication: https://developers.openai.com/apps-sdk/build/auth
- Microsoft identity scopes: https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc
