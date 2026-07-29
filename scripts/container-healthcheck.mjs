#!/usr/bin/env node

import { spawnSync } from "node:child_process";

const officePython = process.env.ONEDRIVE_OFFICE_PYTHON || "/opt/onedrive-office/bin/python3";
const office = spawnSync(
  officePython,
  ["-c", "import docx, openpyxl, pptx, PIL"],
  { stdio: "ignore", timeout: 4_000 }
);
if (office.status !== 0) process.exit(1);

const urls = ["http://127.0.0.1:8765/readyz"];
if (/^(1|true)$/i.test(process.env.ONEDRIVE_OAUTH_COMPAT_ENABLED || "")) {
  urls.push(`http://127.0.0.1:${process.env.ONEDRIVE_OAUTH_COMPAT_PORT || 3010}/healthz`);
}

try {
  const responses = await Promise.all(
    urls.map((url) => fetch(url, { signal: AbortSignal.timeout(4_000) }))
  );
  if (responses.some((response) => !response.ok)) process.exit(1);
} catch {
  process.exit(1);
}
