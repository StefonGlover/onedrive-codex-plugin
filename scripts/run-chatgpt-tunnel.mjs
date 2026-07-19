#!/usr/bin/env node

import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { spawn } from "node:child_process";

const envFile = process.env.ONEDRIVE_TUNNEL_ENV_FILE || join(homedir(), ".config", "tunnel-client", "onedrive-chatgpt.env");
const profileDir = process.env.ONEDRIVE_TUNNEL_PROFILE_DIR || join(homedir(), ".config", "tunnel-client");
const profile = process.env.ONEDRIVE_TUNNEL_PROFILE || "onedrive-chatgpt";
const tunnelClient = process.env.ONEDRIVE_TUNNEL_CLIENT || "/opt/homebrew/bin/tunnel-client";

function envValue(text, name) {
  const match = text.match(new RegExp(`^\\s*${name}\\s*=\\s*(.*?)\\s*$`, "m"));
  if (!match) return "";
  const raw = match[1];
  if (raw.startsWith('"') && raw.endsWith('"')) {
    try {
      return JSON.parse(raw);
    } catch {
      return raw.slice(1, -1);
    }
  }
  if (raw.startsWith("'") && raw.endsWith("'")) return raw.slice(1, -1);
  return raw;
}

let localEnv;
try {
  localEnv = readFileSync(envFile, "utf8");
} catch (error) {
  console.error(`Could not read tunnel credential file: ${error.message}`);
  process.exit(1);
}

const runtimeKey = envValue(localEnv, "CONTROL_PLANE_API_KEY") || envValue(localEnv, "OPENAI_API_KEY");
if (!runtimeKey) {
  console.error("The tunnel credential file does not contain CONTROL_PLANE_API_KEY or OPENAI_API_KEY.");
  process.exit(1);
}

const child = spawn(tunnelClient, ["run", "--profile-dir", profileDir, "--profile", profile], {
  env: { ...process.env, CONTROL_PLANE_API_KEY: runtimeKey },
  stdio: "inherit"
});

for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => child.kill(signal));
}

child.on("error", (error) => {
  console.error(`Could not launch tunnel-client: ${error.message}`);
  process.exit(1);
});

child.on("exit", (code, signal) => {
  if (signal) process.kill(process.pid, signal);
  else process.exit(code ?? 1);
});
