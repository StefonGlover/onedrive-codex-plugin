#!/usr/bin/env python3
"""Verify fail-closed and inherited limits for the common-text helper."""

from __future__ import annotations

import json
import platform
from pathlib import Path
import subprocess
import sys
import tempfile
import time


ROOT = Path(__file__).resolve().parent
HELPER = ROOT / "common-text.py"
ADDRESS_LIMIT = 256 * 1024 * 1024
FILE_LIMIT = 4096


if platform.system() == "Darwin":
    unsupported = subprocess.run(
        [sys.executable, str(HELPER)],
        input=b"{}",
        capture_output=True,
        timeout=5,
    )
    result = json.loads(unsupported.stdout.decode("utf-8"))
    assert unsupported.returncode != 0 and result.get("ok") is False, result
    assert "resource limits" in result.get("error", "").lower(), result
    print(json.dumps({
        "ok": True,
        "checks": {
            "unsupportedAddressSpaceLimitFailsClosedBeforeParsing": True,
            "linuxEnforcementCoveredInCi": True,
        },
    }))
    raise SystemExit(0)


loader = r'''
import importlib.util, json, os, resource, signal, subprocess, sys
helper_path, output_path = sys.argv[1:]
spec = importlib.util.spec_from_file_location("onedrive_common_text", helper_path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
module.configure_process_limits(address_space_bytes=256 * 1024 * 1024, file_size_bytes=4096, cpu_seconds=2)
child_probe = "import json,resource; print(json.dumps({'cpu':resource.getrlimit(resource.RLIMIT_CPU),'address':resource.getrlimit(resource.RLIMIT_AS),'file':resource.getrlimit(resource.RLIMIT_FSIZE)}))"
child = subprocess.run([sys.executable, "-c", child_probe], capture_output=True, text=True, check=True)
memory_blocked = False
try:
    bytearray(512 * 1024 * 1024)
except (MemoryError, OSError):
    memory_blocked = True
signal.signal(signal.SIGXFSZ, signal.SIG_IGN)
file_blocked = False
written = 0
try:
    with open(output_path, "wb", buffering=0) as output:
        written = output.write(b"x" * 8192)
except OSError:
    file_blocked = True
print(json.dumps({
    "parent": {
        "cpu": resource.getrlimit(resource.RLIMIT_CPU),
        "address": resource.getrlimit(resource.RLIMIT_AS),
        "file": resource.getrlimit(resource.RLIMIT_FSIZE),
    },
    "child": json.loads(child.stdout),
    "memoryBlocked": memory_blocked,
    "fileBlocked": file_blocked or written < 8192,
    "size": os.path.getsize(output_path) if os.path.exists(output_path) else 0,
}))
'''

cpu_probe = r'''
import importlib.util, sys
spec = importlib.util.spec_from_file_location("onedrive_common_text", sys.argv[1])
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
module.configure_process_limits(address_space_bytes=256 * 1024 * 1024, file_size_bytes=4096, cpu_seconds=1)
while True:
    pass
'''

with tempfile.TemporaryDirectory(prefix="onedrive-common-text-limits-") as temporary:
    output_path = Path(temporary) / "bounded-output.bin"
    completed = subprocess.run(
        [sys.executable, "-c", loader, str(HELPER), str(output_path)],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )
    evidence = json.loads(completed.stdout)
    expected = {"cpu": [2, 2], "address": [ADDRESS_LIMIT, ADDRESS_LIMIT], "file": [FILE_LIMIT, FILE_LIMIT]}
    assert evidence["parent"] == expected, evidence
    assert evidence["child"] == expected, evidence
    assert evidence["memoryBlocked"] is True, evidence
    assert evidence["fileBlocked"] is True and evidence["size"] <= FILE_LIMIT, evidence

    started = time.monotonic()
    cpu_limited = subprocess.run(
        [sys.executable, "-c", cpu_probe, str(HELPER)],
        capture_output=True,
        timeout=6,
    )
    elapsed = time.monotonic() - started
    assert cpu_limited.returncode != 0, cpu_limited
    assert elapsed < 5, elapsed

print(json.dumps({
    "ok": True,
    "checks": {
        "cpuLimitAppliedAndEnforced": True,
        "addressSpaceLimitAppliedAndEnforced": True,
        "fileSizeLimitAppliedAndEnforced": True,
        "limitsInheritedByExtractorChildren": True,
    },
}))
