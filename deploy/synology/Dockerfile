FROM debian:bookworm-slim AS tunnel-client

ARG TUNNEL_CLIENT_VERSION=0.0.10
ARG TARGETARCH=amd64

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl unzip \
    && rm -rf /var/lib/apt/lists/*

RUN set -eu; \
    case "${TARGETARCH}" in amd64|arm64) ;; *) echo "Unsupported tunnel-client architecture: ${TARGETARCH}" >&2; exit 1 ;; esac; \
    archive="tunnel-client-v${TUNNEL_CLIENT_VERSION}-linux-${TARGETARCH}.zip"; \
    release="https://github.com/openai/tunnel-client/releases/download/v${TUNNEL_CLIENT_VERSION}"; \
    mkdir -p /tmp/tunnel; \
    curl -fsSL "${release}/${archive}" -o "/tmp/tunnel/${archive}"; \
    curl -fsSL "${release}/SHA256SUMS.txt" -o /tmp/tunnel/SHA256SUMS.txt; \
    cd /tmp/tunnel; \
    grep " ${archive}$" SHA256SUMS.txt | sha256sum -c -; \
    unzip -q "${archive}"; \
    install -m 0755 tunnel-client /usr/local/bin/tunnel-client

FROM node:20-bookworm-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates catdoc gosu poppler-utils python3 tesseract-ocr tini \
    && rm -rf /var/lib/apt/lists/*

COPY --from=tunnel-client /usr/local/bin/tunnel-client /usr/local/bin/tunnel-client

WORKDIR /app
COPY . /app
RUN chmod 0755 /app/deploy/synology/entrypoint.sh /app/scripts/run-chatgpt-tunnel.mjs /app/scripts/run-chatgpt-oauth-tunnel.mjs /app/mcp/server.mjs /app/mcp/http-server.mjs

ENV HOME=/home/node \
    ONEDRIVE_OFFICE_PYTHON=/usr/bin/python3 \
    ONEDRIVE_STORAGE_ROOT=/data \
    ONEDRIVE_CACHE_ROOT=/data/cache \
    ONEDRIVE_TUNNEL_CLIENT=/usr/local/bin/tunnel-client \
    ONEDRIVE_TUNNEL_ENV_FILE=/run/onedrive-runtime/tunnel.env \
    ONEDRIVE_TUNNEL_PROFILE_DIR=/data/tunnel \
    ONEDRIVE_TUNNEL_PROFILE=onedrive-chatgpt \
    ONEDRIVE_MCP_HTTP_HOST=127.0.0.1 \
    ONEDRIVE_MCP_HTTP_PORT=3001 \
    ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET_FILE=/run/onedrive-runtime/oauth-api-client.secret

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
  CMD ["node", "-e", "fetch('http://127.0.0.1:8765/readyz').then(r=>{if(!r.ok)process.exit(1)}).catch(()=>process.exit(1))"]

ENTRYPOINT ["/usr/bin/tini", "--", "/app/deploy/synology/entrypoint.sh"]
CMD ["node", "/app/scripts/run-chatgpt-oauth-tunnel.mjs"]
