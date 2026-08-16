#!/bin/sh
set -eu

source_dir=/run/onedrive-source
runtime_dir=/run/onedrive-runtime
profile_dir=${ONEDRIVE_TUNNEL_PROFILE_DIR:-/data/tunnel}
profile_name=${ONEDRIVE_TUNNEL_PROFILE:-onedrive-chatgpt}
token_file=${ONEDRIVE_TOKEN_FILE:-/data/auth/tokens.enc}

if [ ! -f "$source_dir/tunnel.env" ] || [ ! -f "$source_dir/auth-vault.key" ]; then
  echo "The NAS runtime directory must contain tunnel.env and auth-vault.key." >&2
  exit 1
fi
if [ -z "${ONEDRIVE_TUNNEL_ID:-}" ]; then
  echo "ONEDRIVE_TUNNEL_ID is required." >&2
  exit 1
fi

chmod 0700 "$source_dir"
chmod 0600 "$source_dir/tunnel.env" "$source_dir/auth-vault.key"
install -d -m 0700 -o node -g node "$runtime_dir" "$profile_dir" /data/auth /data/cache /data/audit /data/backups /data/chatgpt-uploads /data/downloads /data/materialized-resources /data/office-editing /data/updates /data/watches /data/workspaces
install -m 0600 -o node -g node "$source_dir/tunnel.env" "$runtime_dir/tunnel.env"
install -m 0600 -o node -g node "$source_dir/auth-vault.key" "$runtime_dir/auth-vault.key"
if [ "${ONEDRIVE_MCP_AUTH_MODE:-noauth}" = "oauth" ]; then
  if [ ! -f "$source_dir/oauth-api-client.secret" ] || [ -L "$source_dir/oauth-api-client.secret" ]; then
    echo "OAuth mode requires a regular $source_dir/oauth-api-client.secret file." >&2
    exit 1
  fi
  chmod 0600 "$source_dir/oauth-api-client.secret"
  install -m 0600 -o node -g node "$source_dir/oauth-api-client.secret" "$runtime_dir/oauth-api-client.secret"
fi
case "${ONEDRIVE_OAUTH_COMPAT_ENABLED:-false}" in
  1|[Tt][Rr][Uu][Ee])
    if [ "${ONEDRIVE_MCP_AUTH_MODE:-noauth}" != "oauth" ]; then
      echo "The OAuth compatibility service requires ONEDRIVE_MCP_AUTH_MODE=oauth." >&2
      exit 1
    fi
    if [ ! -f "$source_dir/oauth-chatgpt-client.secret" ] || [ -L "$source_dir/oauth-chatgpt-client.secret" ]; then
      echo "The OAuth compatibility service requires a regular $source_dir/oauth-chatgpt-client.secret file." >&2
      exit 1
    fi
    if cmp -s "$source_dir/oauth-api-client.secret" "$source_dir/oauth-chatgpt-client.secret"; then
      echo "The MCP API and ChatGPT OAuth client registrations must use distinct secret files." >&2
      exit 1
    fi
    chmod 0600 "$source_dir/oauth-chatgpt-client.secret"
    install -m 0600 -o node -g node "$source_dir/oauth-chatgpt-client.secret" "$runtime_dir/oauth-chatgpt-client.secret"
    ;;
esac
if [ -e "$token_file" ] || [ -L "$token_file" ]; then
  if [ ! -f "$token_file" ] || [ -L "$token_file" ]; then
    echo "The encrypted OneDrive token file must be a regular file and must not be a symbolic link." >&2
    exit 1
  fi
  chmod 0600 "$token_file"
  chown node:node "$token_file"
fi
if [ "${ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD:-client_secret_post}" = "none" ]; then
  refresh_store_file=${ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE:-}
  case "$refresh_store_file" in
    /data/auth/*) ;;
    *)
      echo "Public OAuth mode requires its refresh store under /data/auth." >&2
      exit 1
      ;;
  esac
  if [ -e "$refresh_store_file" ] || [ -L "$refresh_store_file" ]; then
    if [ ! -f "$refresh_store_file" ] || [ -L "$refresh_store_file" ]; then
      echo "The OAuth refresh store must be a regular file and must not be a symbolic link." >&2
      exit 1
    fi
    chmod 0600 "$refresh_store_file"
    chown node:node "$refresh_store_file"
  fi
fi

profile_path="$profile_dir/$profile_name.yaml"
cat >"$profile_path" <<EOF
config_version: 1
control_plane:
  base_url: "https://api.openai.com"
  tunnel_id: "$ONEDRIVE_TUNNEL_ID"
  api_key: "env:CONTROL_PLANE_API_KEY"
health:
  listen_addr: "127.0.0.1:8765"
admin_ui:
  open_browser: false
log:
  level: info
  format: json
mcp:
  server_urls:
    - channel: main
      url: "http://127.0.0.1:${ONEDRIVE_MCP_HTTP_PORT:-3001}/mcp"
EOF
chmod 0600 "$profile_path"
chown node:node "$profile_path"

exec gosu node "$@"
