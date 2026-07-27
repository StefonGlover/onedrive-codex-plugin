# Synology DS923+ deployment

The base project runs the OneDrive MCP server and OpenAI Secure MCP Tunnel as an outbound-only Container Manager service. The optional ChatGPT Work OAuth deployment additionally publishes its narrowly scoped compatibility service on NAS loopback only; a trusted HTTPS reverse proxy exposes that OAuth origin without exposing the MCP transport.

The DSM project directory has this layout:

```text
onedrive-chatgpt/
  compose.yaml
  app/       # packaged plugin source
  data/      # persistent cache, backups, audit, and encrypted Microsoft token
  runtime/   # owner-only tunnel.env, auth-vault.key, and OAuth client secrets
```

`runtime/tunnel.env` contains the tunnel runtime API key as `CONTROL_PLANE_API_KEY=...`. `runtime/auth-vault.key` contains a base64-encoded 32-byte encryption key. Neither file belongs in source control or a plugin package.

The entrypoint restricts runtime credentials to mode `0600`, copies them into a private in-memory runtime directory, creates owner-only persistent directories, creates an HTTP-target tunnel profile, and drops from root to the unprivileged `node` account before starting the service. Persistent device-code tokens and the public facade's upstream refresh tokens are encrypted with AES-256-GCM and written atomically under `data/auth`; Work refresh tokens are never stored in plaintext or returned to ChatGPT.

After the project is healthy, run `onedrive_auth_device_start`, complete Microsoft device-code login, then run `onedrive_auth_device_poll`. Re-run a read-only health check and ChatGPT smoke test before stopping the previous tunnel client.

## ChatGPT Work OAuth deployment

The base `compose.yaml` remains `noauth` so upgrading the image cannot strand the verified Chat connection before Entra is configured. To enable Work:

1. Configure the two Entra registrations and `access_as_user` scope described in the root `README.md`. Add the ChatGPT client ID to both the API scope's preauthorized applications and the API app's `knownClientApplications`. The API app must also declare the delegated Microsoft Graph permissions it uses.
2. Copy `compose.oauth.example.yaml` to a private deployment override and replace every placeholder. Keep the values in these three groups distinct:
   - Entra API identity: `ONEDRIVE_MCP_OAUTH_API_RESOURCE` and `ONEDRIVE_MCP_OAUTH_API_SCOPE` use the API registration's `api://<MCP_API_CLIENT_ID>` identifier; `ONEDRIVE_MCP_OAUTH_AUDIENCE` is the API client ID.
   - Public MCP identity: `ONEDRIVE_MCP_PROTECTED_RESOURCE` is the exact externally visible HTTPS MCP endpoint, and `ONEDRIVE_MCP_RESOURCE_METADATA_URL` is the exact public protected-resource metadata URL advertised for this tunnel. Neither may be an `api://` identifier or loopback URL.
   - Public OAuth identity: `ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER` and `ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER` are the same public HTTPS origin for the compatibility service. In public-client mode, register and configure the exact `<issuer>/callback` URI through `ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI`. Keep `ONEDRIVE_MCP_OAUTH_AUTHORITY` pointed at Microsoft so bearer-token issuer/JWKS validation still uses Entra.
   Keep `ONEDRIVE_OAUTH_COMPAT_SCOPES` equal to ChatGPT's narrow `access_as_user offline_access` set. Set `ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES` to the same API resource's `/.default offline_access` set so Entra performs its documented combined-consent flow for the known client and middle-tier API. The adapter rejects any other outer-to-upstream scope substitution.
   Set `ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS` to the separate ChatGPT Entra client application ID. Do not put a client secret in Compose.
3. Save the OneDrive MCP API registration's client secret as `runtime/oauth-api-client.secret`, and save the facade's separate Entra client secret as `runtime/oauth-chatgpt-client.secret`. Set both files to mode `0600`; the two Entra applications must not share a secret. ChatGPT never receives the latter secret in preferred public-PKCE mode. In the explicitly enabled confidential compatibility mode below, configure that same client secret in the ChatGPT app for `client_secret_post`.
4. Deploy with both Compose files so the override sets `ONEDRIVE_MCP_AUTH_MODE=oauth` and `ONEDRIVE_OAUTH_COMPAT_ENABLED=true`.
5. Expose host loopback port `3010` through a trusted public HTTPS reverse proxy. For Tailscale Funnel on this NAS, use an allowed HTTPS port and keep the proxy target on loopback:

   ```sh
   tailscale funnel --bg --https=8443 http://127.0.0.1:3010
   ```

   The resulting issuer must exactly match `ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER`. This deliberately makes only OAuth metadata, `/authorize`, `/callback`, `/token`, and `/healthz` public; the MCP server stays private behind Secure MCP Tunnel. Funnel reaches the service from a shared loopback address, so the adapter never treats the socket address or forwarded headers as a caller identity. Cheap metadata and health requests have no persistent shared-IP lockout. Token POSTs first pass a bounded in-flight gate, body limit, and timeout, then use a capped LRU limiter keyed by a domain-separated hash of the validated one-time facade code or refresh handle. Authorization requests additionally use the validated outer-state/PKCE pair plus a continuously refilling route budget whose ten-minute count and worst-case serialized-byte envelope stay below the ephemeral vault quota.
6. Confirm the container health, public OAuth metadata, and tunnel readiness. Secure MCP Tunnel forwards OAuth discovery and rewrites protected-resource metadata `resource` values and `WWW-Authenticate` `resource_metadata` values to public tunnel-service URLs for the same tunnel ID. Run `tunnel-client doctor --profile onedrive-chatgpt --explain` and verify the final advertised public values before cutover.
7. For a ChatGPT user-defined public client, configure OAuth with `<public-compat-issuer>/authorize` and `<public-compat-issuer>/token`, token authentication method `none`, the API `access_as_user` scope, and `offline_access`. Set `ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE=true` only for the verified ChatGPT Work client shape: Work currently omits an outer PKCE pair, so the adapter generates and enforces an independent S256 pair for the Microsoft leg. If a client does send S256, the adapter verifies that outer pair normally. Register `<public-compat-issuer>/callback` on the Entra client; ChatGPT's own fixed callback remains `ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI`. Keep the encrypted store at `/data/auth/oauth-compat-refresh-handles.json` and derive its domain-separated key from `/run/onedrive-runtime/auth-vault.key`. The adapter persists only hashed one-time handles and authenticated ciphertext. OAuth client-secret rotation preserves active sessions; auth-vault key rotation intentionally requires reconnecting after the old store is securely retired. Do not request `openid` or `profile`; this compatibility tier is OAuth-only and does not issue ID tokens. Connect and run the read-only OneDrive health check in Work.

If a different ChatGPT client cannot use the public facade, switch only that
deployment to the opt-in confidential lane:

```yaml
ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "client_secret_post"
ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "true"
```

Configure the same client secret and `client_secret_post` in ChatGPT, and
register the exact ChatGPT callback as an Entra Web redirect. This mode still
validates the exact client, callback, resource, scopes, and state, then requires
the secret in constant time at `/token`. The provider callback and encrypted
facade refresh store are unused. Leave the compatibility flag false for every
other deployment. Prefer the public lane with generated provider PKCE for the
verified Work request shape.

The container starts `mcp/http-server.mjs` on `127.0.0.1:3001` and points Secure MCP Tunnel at that private transport URL. The optional compatibility service listens separately on container port `3010`, mapped only to NAS loopback. In public mode it substitutes its own provider callback/state, forces Microsoft's callback to bounded query mode, rejects callback POST bodies, maps bounded provider codes to one-time facade codes, verifies any outer S256 pair, generates an independent provider S256 pair for the explicitly allowed Work no-PKCE shape, and rotates encrypted one-time refresh handles; direct provider codes and upstream refresh tokens never reach ChatGPT. Outer state is capped at 512 UTF-8 bytes and provider codes at 3 KiB. The shared encrypted file enforces independent 320 KiB refresh and 640 KiB ephemeral partitions within a 1 MiB total cap, so forged authorization traffic cannot consume the space reserved for refresh rotation. Compose rotates `json-file` logs at 10 MiB with three files; successful health diagnostics are suppressed and repeated 429 diagnostics are sampled once per minute. The HTTP MCP server independently validates the resulting bearer token's issuer, API audience, delegated scope, and authorized client application, then uses the validated tenant for OBO Graph exchange.
