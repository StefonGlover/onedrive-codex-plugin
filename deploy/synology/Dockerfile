FROM golang:1.26.6-bookworm@sha256:116d58cbd88c1297624acc6e967a060012422bacf9930927e23fb719189c6f36 AS tunnel-client

ARG TUNNEL_CLIENT_VERSION=0.0.12
ARG TUNNEL_CLIENT_SOURCE_SHA256=aeb0a91f00515f77cc25d23de5e1ea4d934a57476e6087bf85c6f527064532eb
ARG TUNNEL_CLIENT_SOURCE_COMMIT=881c9a8fed7cccbe6607cd419863bbca506b8215
ARG TUNNEL_CLIENT_X_NET_VERSION=v0.56.0
ARG TUNNEL_CLIENT_OTEL_VERSION=v1.43.0
ARG TUNNEL_CLIENT_OTEL_CONTRIB_VERSION=v0.65.0
ARG TARGETARCH=amd64

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && rm -rf /var/lib/apt/lists/*

RUN set -eu; \
    test "${TUNNEL_CLIENT_VERSION}" = "0.0.12"; \
    case "${TARGETARCH}" in amd64|arm64) ;; *) echo "Unsupported tunnel-client architecture: ${TARGETARCH}" >&2; exit 1 ;; esac; \
    archive="tunnel-client-${TUNNEL_CLIENT_SOURCE_COMMIT}.tar.gz"; \
    source_url="https://codeload.github.com/openai/tunnel-client/tar.gz/${TUNNEL_CLIENT_SOURCE_COMMIT}"; \
    mkdir -p /tmp/tunnel; \
    curl -fsSL "${source_url}" -o "/tmp/tunnel/${archive}"; \
    cd /tmp/tunnel; \
    printf '%s  %s\n' "${TUNNEL_CLIENT_SOURCE_SHA256}" "${archive}" | sha256sum -c -; \
    tar -xzf "${archive}"; \
    mv "tunnel-client-${TUNNEL_CLIENT_SOURCE_COMMIT}" source; \
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
    module="github.com/openai/tunnel-client"; \
    build_flags="-trimpath -buildvcs=false"; \
    go_version="$(go env GOVERSION)"; \
    ldflags="-X ${module}/pkg/version.semanticVersion=${TUNNEL_CLIENT_VERSION} -X ${module}/pkg/version.GitSHA=${TUNNEL_CLIENT_SOURCE_COMMIT} -X ${module}/pkg/version.GoVersion=${go_version} -X '${module}/pkg/version.BuildFlags=${build_flags}' -X ${module}/pkg/version.Flavor=runtime"; \
    env GOWORK=off GOOS=linux GOARCH="${TARGETARCH}" CGO_ENABLED=0 \
      go build -mod=readonly -trimpath -buildvcs=false -ldflags "${ldflags}" \
      -o /usr/local/bin/tunnel-client ./cmd/client-runtime; \
    chmod 0755 /usr/local/bin/tunnel-client; \
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
