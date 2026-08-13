#!/usr/bin/env python3
"""Exercise CPU, address-space, and output-file limits for the renderer launcher."""

from __future__ import annotations

import json
import platform
from pathlib import Path
import signal
import subprocess
import sys
import tempfile
import time


ROOT = Path(__file__).resolve().parent
LAUNCHER = ROOT / "pdftoppm-limited.py"
ADDRESS_LIMIT = 256 * 1024 * 1024
FILE_LIMIT = 4096


def launcher_command(cpu_seconds: int, command: list[str]) -> list[str]:
    return [
        sys.executable,
        str(LAUNCHER),
        "--cpu-seconds",
        str(cpu_seconds),
        "--address-space-bytes",
        str(ADDRESS_LIMIT),
        "--file-size-bytes",
        str(FILE_LIMIT),
        "--",
        *command,
    ]


if platform.system() == "Darwin":
    unsupported = subprocess.run(
        launcher_command(1, ["/usr/bin/true"]),
        capture_output=True,
        text=True,
        timeout=5,
    )
    assert unsupported.returncode == 70, unsupported
    assert "resource limits could not be applied" in unsupported.stderr, unsupported.stderr
    print(json.dumps({
        "ok": True,
        "checks": {
            "unsupportedAddressSpaceLimitFailsClosed": True,
            "linuxEnforcementCoveredInCi": True,
        },
    }))
    raise SystemExit(0)


with tempfile.TemporaryDirectory(prefix="onedrive-pdftoppm-limits-") as temp_dir:
    output_path = Path(temp_dir) / "bounded-output.bin"
    probe = r'''
import json, os, resource, signal, sys
target = sys.argv[1]
memory_blocked = False
try:
    bytearray(512 * 1024 * 1024)
except (MemoryError, OSError):
    memory_blocked = True
signal.signal(signal.SIGXFSZ, signal.SIG_IGN)
file_blocked = False
written = 0
try:
    with open(target, "wb", buffering=0) as output:
        written = output.write(b"x" * 8192)
except OSError:
    file_blocked = True
size = os.path.getsize(target) if os.path.exists(target) else 0
print(json.dumps({
    "cpu": resource.getrlimit(resource.RLIMIT_CPU),
    "address": resource.getrlimit(resource.RLIMIT_AS),
    "file": resource.getrlimit(resource.RLIMIT_FSIZE),
    "memoryBlocked": memory_blocked,
    "fileBlocked": file_blocked or written < 8192,
    "size": size,
}))
'''
    completed = subprocess.run(
        launcher_command(2, [sys.executable, "-c", probe, str(output_path)]),
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )
    evidence = json.loads(completed.stdout)
    assert evidence["cpu"] == [2, 2], evidence
    assert evidence["address"] == [ADDRESS_LIMIT, ADDRESS_LIMIT], evidence
    assert evidence["file"] == [FILE_LIMIT, FILE_LIMIT], evidence
    assert evidence["memoryBlocked"] is True, evidence
    assert evidence["fileBlocked"] is True and evidence["size"] <= FILE_LIMIT, evidence

    started = time.monotonic()
    cpu_limited = subprocess.run(
        launcher_command(1, [sys.executable, "-c", "while True: pass"]),
        capture_output=True,
        timeout=6,
    )
    elapsed = time.monotonic() - started
    assert cpu_limited.returncode != 0, cpu_limited
    assert elapsed < 5, elapsed

    invalid = subprocess.run(
        [
            sys.executable,
            str(LAUNCHER),
            "--cpu-seconds",
            "21",
            "--address-space-bytes",
            str(ADDRESS_LIMIT),
            "--file-size-bytes",
            str(FILE_LIMIT),
            "--",
            sys.executable,
            "-c",
            "pass",
        ],
        capture_output=True,
        text=True,
        timeout=5,
    )
    assert invalid.returncode != 0 and "outside the permitted range" in invalid.stderr, invalid.stderr

print(json.dumps({
    "ok": True,
    "checks": {
        "cpuLimitAppliedAndEnforced": True,
        "addressSpaceLimitAppliedAndEnforced": True,
        "fileSizeLimitAppliedAndEnforced": True,
        "outOfRangeLimitRejected": True,
    },
}))
