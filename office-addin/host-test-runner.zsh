#!/bin/zsh

set -euo pipefail
unsetopt BG_NICE

if [[ $# -ne 1 || ! "$1" =~ '^[A-Za-z0-9]([A-Za-z0-9._-]{0,62}[A-Za-z0-9])?$' ]]; then
  print -u2 "Usage: office-addin/host-test-runner.zsh <1-64 character run-id>"
  exit 2
fi
if [[ "$(uname -s)" != "Darwin" ]]; then
  print -u2 "The real Office host runner is supported only on macOS."
  exit 2
fi

readonly RUN_ID="$1"
readonly OFFICE_ROOT="${0:A:h}"
readonly TEST_DIR="/tmp/codex-onedrive-office-${RUN_ID}"
readonly CA_NAME="Codex OneDrive Office Companion Root ${RUN_ID}"
readonly LEAF_NAME="Codex OneDrive Office Companion Loopback ${RUN_ID}"
readonly CA_CERT_PATH="${TEST_DIR}/ca-cert.pem"
readonly CA_KEY_PATH="${TEST_DIR}/ca-key.pem"
readonly CERT_PATH="${TEST_DIR}/cert.pem"
readonly KEY_PATH="${TEST_DIR}/key.pem"
readonly CSR_PATH="${TEST_DIR}/cert.csr"
readonly EXT_PATH="${TEST_DIR}/cert-ext.cnf"
readonly KEYCHAIN="${HOME}/Library/Keychains/login.keychain-db"
readonly MANIFEST_NAME="codex-onedrive-office-${RUN_ID}.xml"
readonly TEMP_MANIFEST="${TEST_DIR}/${MANIFEST_NAME}"
readonly SERVER_LOG="${TEST_DIR}/server.log"
readonly -a OFFICE_CONTAINERS=(com.microsoft.Word com.microsoft.Excel com.microsoft.Powerpoint)

typeset -a MANIFEST_PATHS=()
typeset -a CREATED_WEF_DIRS=()
typeset SERVER_PID=""
typeset CA_SHA256=""
typeset LEAF_SHA256=""
typeset MANIFEST_ID=""
typeset CLEANUP_STARTED=0

report_request_evidence() {
  typeset total_count
  total_count="$(grep -cF '"event":"office-companion-request"' "$SERVER_LOG" 2>/dev/null || true)"
  print "OFFICE_HOST_REQUEST_EVIDENCE runId=${RUN_ID} total=${total_count:-0}"
  typeset asset_path
  typeset ok_count
  for asset_path in \
    /office-addin/taskpane.html \
    /office-addin/taskpane.js \
    /office-addin/icon-16.png \
    /office-addin/icon-32.png \
    /office-addin/icon-64.png \
    /office-addin/icon-80.png; do
    ok_count="$(grep -F "\"path\":\"${asset_path}\"" "$SERVER_LOG" 2>/dev/null | grep -cF '"status":200' || true)"
    print "OFFICE_HOST_REQUEST_ASSET runId=${RUN_ID} path=${asset_path} status200=${ok_count:-0}"
  done
}

for command_name in node openssl security curl lsof ps; do
  if ! command -v "$command_name" >/dev/null 2>&1; then
    print -u2 "Required command is unavailable: ${command_name}"
    exit 2
  fi
done
if [[ -e "$TEST_DIR" ]]; then
  print -u2 "Run-scoped test directory already exists; refusing to reuse it: ${TEST_DIR}"
  exit 2
fi
if lsof -nP -iTCP:3443 -sTCP:LISTEN >/dev/null 2>&1; then
  print -u2 "TCP port 3443 is already in use; refusing to disturb the existing listener."
  exit 2
fi

cleanup() {
  if (( CLEANUP_STARTED )); then return 0; fi
  CLEANUP_STARTED=1
  set +e
  typeset cleanup_status=0

  for manifest_path in "${MANIFEST_PATHS[@]}"; do
    rm -f -- "$manifest_path"
    if [[ -e "$manifest_path" ]]; then
      print -u2 "Cleanup could not remove run manifest: ${manifest_path}"
      cleanup_status=1
    fi
  done

  if [[ -n "$SERVER_PID" ]] && kill -0 "$SERVER_PID" >/dev/null 2>&1; then
    typeset server_command
    server_command="$(ps -p "$SERVER_PID" -o command= 2>/dev/null)"
    if [[ "$server_command" == *"${OFFICE_ROOT}/serve.mjs"* && "$server_command" == *"${CERT_PATH}"* ]]; then
      kill "$SERVER_PID" >/dev/null 2>&1
      wait "$SERVER_PID" >/dev/null 2>&1
    else
      print -u2 "Cleanup refused to stop PID ${SERVER_PID} because it no longer matches this run."
      cleanup_status=1
    fi
  fi

  if [[ -n "$LEAF_SHA256" ]]; then
    typeset leaf_listing
    leaf_listing="$(security find-certificate -Z -c "$LEAF_NAME" "$KEYCHAIN" 2>/dev/null)"
    if [[ "$leaf_listing" == *"${LEAF_SHA256}"* ]]; then
      security delete-certificate -t -Z "$LEAF_SHA256" "$KEYCHAIN" >/dev/null
    fi
    leaf_listing="$(security find-certificate -Z -c "$LEAF_NAME" "$KEYCHAIN" 2>/dev/null)"
    if [[ "$leaf_listing" == *"${LEAF_SHA256}"* ]]; then
      print -u2 "Cleanup found an unexpected imported loopback leaf and could not remove it: ${LEAF_SHA256}"
      cleanup_status=1
    fi
  fi

  if [[ -n "$CA_SHA256" ]]; then
    typeset certificate_listing
    certificate_listing="$(security find-certificate -Z -c "$CA_NAME" "$KEYCHAIN" 2>/dev/null)"
    if [[ "$certificate_listing" == *"${CA_SHA256}"* ]]; then
      security delete-certificate -t -Z "$CA_SHA256" "$KEYCHAIN" >/dev/null
    fi
    certificate_listing="$(security find-certificate -Z -c "$CA_NAME" "$KEYCHAIN" 2>/dev/null)"
    if [[ "$certificate_listing" == *"${CA_SHA256}"* ]]; then
      security remove-trusted-cert "$CA_CERT_PATH" >/dev/null 2>&1 || true
      security delete-certificate -t -Z "$CA_SHA256" "$KEYCHAIN" >/dev/null 2>&1 || true
      certificate_listing="$(security find-certificate -Z -c "$CA_NAME" "$KEYCHAIN" 2>/dev/null)"
      if [[ "$certificate_listing" == *"${CA_SHA256}"* ]]; then
        print -u2 "Cleanup could not remove the exact certificate authority and its user trust settings: ${CA_SHA256}"
        cleanup_status=1
      fi
    fi
  fi

  if [[ -e "$CERT_PATH" ]] && security verify-cert -q -L -c "$CERT_PATH" -p ssl -n 127.0.0.1 -k "$KEYCHAIN" >/dev/null 2>&1; then
    print -u2 "Cleanup verification found the run-specific loopback leaf still trusted."
    cleanup_status=1
  fi

  for wef_dir in "${(Oa)CREATED_WEF_DIRS[@]}"; do
    rmdir -- "$wef_dir" >/dev/null 2>&1 || true
  done

  if [[ "$TEST_DIR" == "/tmp/codex-onedrive-office-${RUN_ID}" ]]; then
    rm -rf -- "$TEST_DIR"
  else
    print -u2 "Cleanup refused an unexpected test directory: ${TEST_DIR}"
    cleanup_status=1
  fi
  if [[ -e "$TEST_DIR" ]]; then
    print -u2 "Cleanup could not remove the run-scoped test directory: ${TEST_DIR}"
    cleanup_status=1
  fi
  if lsof -nP -iTCP:3443 -sTCP:LISTEN >/dev/null 2>&1; then
    print -u2 "Cleanup verification found a remaining listener on TCP port 3443."
    cleanup_status=1
  fi

  if (( cleanup_status == 0 )); then
    print "OFFICE_HOST_CLEANUP_OK runId=${RUN_ID} manifests=absent certificateAndTrust=absent server=stopped tempDir=absent port3443=closed"
  else
    print -u2 "OFFICE_HOST_CLEANUP_INCOMPLETE runId=${RUN_ID}"
  fi
  return "$cleanup_status"
}

trap cleanup EXIT
trap 'exit 130' HUP INT TERM

mkdir -m 700 -- "$TEST_DIR"
openssl req -x509 -newkey rsa:2048 -sha256 -days 1 -nodes \
  -keyout "$CA_KEY_PATH" -out "$CA_CERT_PATH" \
  -subj "/CN=${CA_NAME}" \
  -addext "basicConstraints=critical,CA:TRUE,pathlen:0" \
  -addext "keyUsage=critical,keyCertSign,cRLSign" \
  -addext "subjectKeyIdentifier=hash" >/dev/null 2>&1
openssl req -new -newkey rsa:2048 -sha256 -nodes \
  -keyout "$KEY_PATH" -out "$CSR_PATH" \
  -subj "/CN=${LEAF_NAME}" >/dev/null 2>&1
{
  print '[loopback_leaf]'
  print 'subjectAltName=IP:127.0.0.1,DNS:localhost'
  print 'basicConstraints=critical,CA:FALSE'
  print 'keyUsage=critical,digitalSignature,keyEncipherment'
  print 'extendedKeyUsage=serverAuth'
  print 'subjectKeyIdentifier=hash'
  print 'authorityKeyIdentifier=keyid,issuer'
} >"$EXT_PATH"
typeset SERIAL_HEX
SERIAL_HEX="$(openssl rand -hex 16)"
if [[ ! "$SERIAL_HEX" =~ '^[0-9A-Fa-f]{32}$' ]]; then
  print -u2 "Could not generate the temporary loopback certificate serial."
  exit 1
fi
openssl x509 -req -in "$CSR_PATH" -CA "$CA_CERT_PATH" -CAkey "$CA_KEY_PATH" \
  -set_serial "0x${SERIAL_HEX}" -days 1 -sha256 \
  -extfile "$EXT_PATH" -extensions loopback_leaf -out "$CERT_PATH" >/dev/null 2>&1
chmod 600 "$CA_KEY_PATH" "$CA_CERT_PATH" "$KEY_PATH" "$CERT_PATH" "$CSR_PATH" "$EXT_PATH"
openssl verify -CAfile "$CA_CERT_PATH" -purpose sslserver "$CERT_PATH" >/dev/null
CA_SHA256="$(openssl x509 -in "$CA_CERT_PATH" -noout -fingerprint -sha256 | awk -F= '{print $2}' | tr -d ':')"
LEAF_SHA256="$(openssl x509 -in "$CERT_PATH" -noout -fingerprint -sha256 | awk -F= '{print $2}' | tr -d ':')"
if [[ ! "$CA_SHA256" =~ '^[0-9A-Fa-f]{64}$' || ! "$LEAF_SHA256" =~ '^[0-9A-Fa-f]{64}$' ]]; then
  print -u2 "Could not compute the temporary certificate fingerprints."
  exit 1
fi

node "${OFFICE_ROOT}/prepare-test-manifest.mjs" --run-id="$RUN_ID" --output="$TEMP_MANIFEST"
MANIFEST_ID="$(sed -n 's#.*<Id>\([^<]*\)</Id>.*#\1#p' "$TEMP_MANIFEST")"
if [[ ! "$MANIFEST_ID" =~ '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[1-5][0-9A-Fa-f]{3}-[89AaBb][0-9A-Fa-f]{3}-[0-9A-Fa-f]{12}$' ]]; then
  print -u2 "Generated Office manifest did not contain one valid UUID."
  exit 1
fi
security add-trusted-cert -r trustRoot -p ssl -s 127.0.0.1 -k "$KEYCHAIN" "$CA_CERT_PATH"
security verify-cert -q -L -c "$CERT_PATH" -p ssl -n 127.0.0.1 -k "$KEYCHAIN"

node "${OFFICE_ROOT}/serve.mjs" --cert "$CERT_PATH" --key "$KEY_PATH" >"$SERVER_LOG" 2>&1 &
SERVER_PID=$!
typeset server_ready=0
for _attempt in {1..50}; do
  if ! kill -0 "$SERVER_PID" >/dev/null 2>&1; then
    print -u2 "Office companion HTTPS server exited during startup."
    tail -20 "$SERVER_LOG" >&2
    exit 1
  fi
  if curl --silent --show-error --fail --cacert "$CA_CERT_PATH" \
    "https://127.0.0.1:3443/office-addin/taskpane.html" >/dev/null; then
    server_ready=1
    break
  fi
  sleep 0.1
done
if (( ! server_ready )); then
  print -u2 "Office companion HTTPS server did not become ready on port 3443."
  exit 1
fi

for container in "${OFFICE_CONTAINERS[@]}"; do
  typeset documents_dir="${HOME}/Library/Containers/${container}/Data/Documents"
  typeset wef_dir="${documents_dir}/wef"
  typeset manifest_path="${wef_dir}/${MANIFEST_ID}.${MANIFEST_NAME}"
  if [[ ! -d "$documents_dir" ]]; then
    print -u2 "Office container Documents directory is missing: ${documents_dir}"
    exit 1
  fi
  if [[ -e "$manifest_path" ]]; then
    print -u2 "Run-specific manifest already exists; refusing to overwrite it: ${manifest_path}"
    exit 1
  fi
  if [[ ! -d "$wef_dir" ]]; then
    CREATED_WEF_DIRS+=("$wef_dir")
    mkdir -m 700 -- "$wef_dir"
  fi
  MANIFEST_PATHS+=("$manifest_path")
  ln -- "$TEMP_MANIFEST" "$manifest_path"
done

print "OFFICE_HOST_SETUP_READY runId=${RUN_ID} serverPid=${SERVER_PID} caSha256=${CA_SHA256} leafSha256=${LEAF_SHA256} manifestId=${MANIFEST_ID} manifest=${MANIFEST_ID}.${MANIFEST_NAME}"
print "The helper will clean up on Return, Control-C, termination, or setup failure. It never clears shared Office caches."
print -n "After closing all three unsaved scratch documents and quitting Office, press Return to clean up: "
read -r

report_request_evidence
trap - EXIT HUP INT TERM
cleanup
