FROM golang:1.26.6-bookworm@sha256:116d58cbd88c1297624acc6e967a060012422bacf9930927e23fb719189c6f36 AS tunnel-client

ARG TUNNEL_CLIENT_VERSION=0.0.12
ARG TUNNEL_CLIENT_SOURCE_ASSET_ID=521784599
ARG TUNNEL_CLIENT_CHECKSUMS_ASSET_ID=521784601
ARG TUNNEL_CLIENT_SOURCE_SHA256=813b300d24fd7bd0115ad4b171972c59399c40593b0fa604e9cf1fc3c795587e
ARG TUNNEL_CLIENT_SOURCE_COMMIT=881c9a8fed7cccbe6607cd419863bbca506b8215
ARG TUNNEL_CLIENT_X_NET_VERSION=v0.56.0
ARG TUNNEL_CLIENT_OTEL_VERSION=v1.43.0
ARG TUNNEL_CLIENT_OTEL_CONTRIB_VERSION=v0.65.0
ARG TARGETARCH=amd64

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl python3 \
    && rm -rf /var/lib/apt/lists/*

RUN set -eu; \
    case "${TARGETARCH}" in amd64|arm64) ;; *) echo "Unsupported tunnel-client architecture: ${TARGETARCH}" >&2; exit 1 ;; esac; \
    archive="tunnel-client-runtime-source-v${TUNNEL_CLIENT_VERSION}.tar.gz"; \
    asset_api="https://api.github.com/repos/openai/tunnel-client/releases/assets"; \
    mkdir -p /tmp/tunnel; \
    curl -fsSL -H 'Accept: application/octet-stream' -H 'X-GitHub-Api-Version: 2022-11-28' "${asset_api}/${TUNNEL_CLIENT_SOURCE_ASSET_ID}" -o "/tmp/tunnel/${archive}"; \
    curl -fsSL -H 'Accept: application/octet-stream' -H 'X-GitHub-Api-Version: 2022-11-28' "${asset_api}/${TUNNEL_CLIENT_CHECKSUMS_ASSET_ID}" -o /tmp/tunnel/SHA256SUMS.txt; \
    cd /tmp/tunnel; \
    grep " ${archive}$" SHA256SUMS.txt | sha256sum -c -; \
    printf '%s  %s\n' "${TUNNEL_CLIENT_SOURCE_SHA256}" "${archive}" | sha256sum -c -; \
    tar -xzf "${archive}"; \
    mv "tunnel-client-runtime-source-v${TUNNEL_CLIENT_VERSION}" source; \
    test "$(python3 -c 'import json; print(json.load(open("source/SOURCE_EXPORT_METADATA/manifest.json"))["sourceCommit"])')" = "${TUNNEL_CLIENT_SOURCE_COMMIT}"; \
    cd source; \
    go mod edit -go=1.26.6; \
    go get \
      "golang.org/x/net@${TUNNEL_CLIENT_X_NET_VERSION}" \
      "go.opentelemetry.io/otel@${TUNNEL_CLIENT_OTEL_VERSION}" \
      "go.opentelemetry.io/otel/metric@${TUNNEL_CLIENT_OTEL_VERSION}" \
      "go.opentelemetry.io/otel/sdk@${TUNNEL_CLIENT_OTEL_VERSION}" \
      "go.opentelemetry.io/otel/sdk/metric@${TUNNEL_CLIENT_OTEL_VERSION}" \
      "go.opentelemetry.io/otel/trace@${TUNNEL_CLIENT_OTEL_VERSION}" \
      "go.opentelemetry.io/otel/exporters/prometheus@${TUNNEL_CLIENT_OTEL_CONTRIB_VERSION}" \
      "go.opentelemetry.io/contrib/instrumentation/net/http/otelhttp@${TUNNEL_CLIENT_OTEL_CONTRIB_VERSION}"; \
    go mod tidy; \
    ./scripts/rebuild_runtime_from_source.sh \
      --flavor runtime \
      --goos linux \
      --goarch "${TARGETARCH}" \
      --output /usr/local/bin/tunnel-client; \
    /usr/local/bin/tunnel-client --version

FROM node:24-bookworm-slim@sha256:3638d9a6fe4030bd716be989438248074489337ba3275657f93595428be4fc03

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates catdoc gosu poppler-utils python3 python3-venv tesseract-ocr tini \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /usr/local/lib/node_modules/npm /usr/local/lib/node_modules/corepack \
    && rm -f /usr/local/bin/npm /usr/local/bin/npx /usr/local/bin/corepack

COPY --from=tunnel-client /usr/local/bin/tunnel-client /usr/local/bin/tunnel-client

COPY scripts/requirements-office-test.txt /tmp/onedrive-office-requirements.txt
RUN python3 -m venv /opt/onedrive-office \
    && /opt/onedrive-office/bin/pip install --no-cache-dir --requirement /tmp/onedrive-office-requirements.txt \
    && /opt/onedrive-office/bin/python3 -c "import docx, openpyxl, pptx, PIL" \
    && /opt/onedrive-office/bin/python3 -m pip uninstall --yes pip setuptools \
    && /opt/onedrive-office/bin/python3 -c "import docx, openpyxl, pptx, PIL" \
    && rm -f /tmp/onedrive-office-requirements.txt

WORKDIR /app
COPY . /app
RUN chmod 0755 /app/deploy/synology/entrypoint.sh /app/scripts/run-chatgpt-tunnel.mjs /app/scripts/run-chatgpt-oauth-tunnel.mjs /app/mcp/server.mjs /app/mcp/http-server.mjs /app/mcp/oauth-compat-server.mjs

ENV HOME=/home/node \
    ONEDRIVE_OFFICE_PYTHON=/opt/onedrive-office/bin/python3 \
    ONEDRIVE_STORAGE_ROOT=/data \
    ONEDRIVE_CACHE_ROOT=/data/cache \
    ONEDRIVE_TUNNEL_CLIENT=/usr/local/bin/tunnel-client \
    ONEDRIVE_TUNNEL_ENV_FILE=/run/onedrive-runtime/tunnel.env \
    ONEDRIVE_TUNNEL_PROFILE_DIR=/data/tunnel \
    ONEDRIVE_TUNNEL_PROFILE=onedrive-chatgpt \
    ONEDRIVE_TOOL_PROFILE=chatgpt \
    ONEDRIVE_MCP_HTTP_HOST=127.0.0.1 \
    ONEDRIVE_MCP_HTTP_PORT=3001 \
    ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET_FILE=/run/onedrive-runtime/oauth-api-client.secret \
    ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE=/run/onedrive-runtime/oauth-chatgpt-client.secret

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
  CMD ["node", "/app/scripts/container-healthcheck.mjs"]

ENTRYPOINT ["/usr/bin/tini", "--", "/app/deploy/synology/entrypoint.sh"]
CMD ["node", "/app/scripts/run-chatgpt-oauth-tunnel.mjs"]
