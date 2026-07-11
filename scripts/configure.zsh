#!/usr/bin/env zsh
set -euo pipefail

config_dir="$HOME/.codex/onedrive-plugin"
config_file="$config_dir/config.json"

mkdir -p "$config_dir"
chmod 700 "$config_dir"

printf "Microsoft app client ID: "
read -r client_id
if [[ -z "$client_id" ]]; then
  echo "Client ID is required." >&2
  exit 1
fi

printf "Tenant [common]: "
read -r tenant
tenant="${tenant:-common}"

printf "Scopes [offline_access User.Read Files.ReadWrite]: "
read -r scopes
scopes="${scopes:-offline_access User.Read Files.ReadWrite}"

printf "Keychain service [Codex OneDrive]: "
read -r keychain_service
keychain_service="${keychain_service:-Codex OneDrive}"

node - "$config_file" "$client_id" "$tenant" "$scopes" "$keychain_service" <<'NODE'
const { readFileSync, writeFileSync } = require("node:fs");
const [path, clientId, tenant, scopes, keychainService] = process.argv.slice(2);
let existing = {};
try {
  existing = JSON.parse(readFileSync(path, "utf8"));
} catch {
  existing = {};
}
writeFileSync(path, JSON.stringify({
  ...existing,
  clientId,
  tenant,
  scopes,
  keychainService
}, null, 2) + "\n", { mode: 0o600 });
NODE

chmod 600 "$config_file"
echo "Wrote $config_file"
echo "Next: reinstall/refresh the plugin, start a fresh Codex thread, then run onedrive_config with checkToken=true. Start device login only if no reusable credential exists."
