# Synology DS923+ deployment

This project runs the OneDrive MCP server and OpenAI Secure MCP Tunnel as an outbound-only Container Manager service. It does not publish an inbound port.

The DSM project directory has this layout:

```text
onedrive-chatgpt/
  compose.yaml
  app/       # packaged plugin source
  data/      # persistent cache, backups, audit, and encrypted Microsoft token
  runtime/   # owner-only tunnel.env and auth-vault.key
```

`runtime/tunnel.env` contains the tunnel runtime API key as `CONTROL_PLANE_API_KEY=...`. `runtime/auth-vault.key` contains a base64-encoded 32-byte encryption key. Neither file belongs in source control or a plugin package.

The entrypoint restricts both runtime files to mode `0600`, copies them into a private in-memory runtime directory, creates the tunnel profile, and drops from root to the unprivileged `node` account before starting the service. Persistent Microsoft tokens are encrypted with AES-256-GCM and written atomically under `data/auth`.

After the project is healthy, run `onedrive_auth_device_start`, complete Microsoft device-code login, then run `onedrive_auth_device_poll`. Re-run a read-only health check and ChatGPT smoke test before stopping the previous tunnel client.
