# Synology DS923+ deployment

This project runs the OneDrive MCP server and OpenAI Secure MCP Tunnel as an outbound-only Container Manager service. It does not publish an inbound port.

The DSM project directory has this layout:

```text
onedrive-chatgpt/
  compose.yaml
  app/       # packaged plugin source
  data/      # persistent cache, backups, audit, and encrypted Microsoft token
  runtime/   # owner-only tunnel.env, auth-vault.key, and optional OAuth API secret
```

`runtime/tunnel.env` contains the tunnel runtime API key as `CONTROL_PLANE_API_KEY=...`. `runtime/auth-vault.key` contains a base64-encoded 32-byte encryption key. Neither file belongs in source control or a plugin package.

The entrypoint restricts runtime credentials to mode `0600`, copies them into a private in-memory runtime directory, creates an HTTP-target tunnel profile, and drops from root to the unprivileged `node` account before starting the service. Persistent device-code tokens are encrypted with AES-256-GCM and written atomically under `data/auth`; delegated Work tokens remain request-scoped and are not persisted.

After the project is healthy, run `onedrive_auth_device_start`, complete Microsoft device-code login, then run `onedrive_auth_device_poll`. Re-run a read-only health check and ChatGPT smoke test before stopping the previous tunnel client.

## ChatGPT Work OAuth deployment

The base `compose.yaml` remains `noauth` so upgrading the image cannot strand the verified Chat connection before Entra is configured. To enable Work:

1. Configure the two Entra registrations and `access_as_user` scope described in the root `README.md`.
2. Copy `compose.oauth.example.yaml` to a private deployment override and replace every placeholder. Do not put a client secret in Compose.
3. Save the OneDrive MCP API registration's client secret as `runtime/oauth-api-client.secret` and set mode `0600`.
4. Deploy with both Compose files so the override sets `ONEDRIVE_MCP_AUTH_MODE=oauth`.
5. Confirm the container health and tunnel readiness. Tunnel discovery should report the protected-resource metadata route and Entra OIDC discovery as healthy.
6. Recreate the ChatGPT app as OAuth, using the separate ChatGPT client registration. Add ChatGPT's exact callback URL to that registration, connect, and run the read-only OneDrive health check in Work.

The container starts `mcp/http-server.mjs` on `127.0.0.1:3001` and points Secure MCP Tunnel at that loopback URL. No host port is published. The HTTP server validates the ChatGPT bearer token, uses OBO to acquire a short-lived Graph token, and scopes local workspace/watch state to an opaque issuer-plus-user identity and the active drive ID.
