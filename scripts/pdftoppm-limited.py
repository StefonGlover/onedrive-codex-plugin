#!/usr/bin/env python3
"""Exec a PDF renderer with fail-closed per-process POSIX resource limits."""

from __future__ import annotations

import os
import resource
import sys


MAX_CPU_SECONDS = 20
MIN_ADDRESS_SPACE_BYTES = 64 * 1024 * 1024
MAX_ADDRESS_SPACE_BYTES = 1024 * 1024 * 1024
MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024 + 4096


def fail(message: str, status: int = 64) -> "None":
    print(f"pdftoppm-limited: {message}", file=sys.stderr)
    raise SystemExit(status)


def bounded_integer(raw: str, label: str, minimum: int, maximum: int) -> int:
    if not raw.isascii() or not raw.isdecimal():
        fail(f"{label} must be a decimal integer")
    value = int(raw)
    if value < minimum or value > maximum:
        fail(f"{label} is outside the permitted range")
    return value


def parse_arguments(argv: list[str]) -> tuple[int, int, int, list[str]]:
    try:
        separator = argv.index("--")
    except ValueError:
        fail("a -- command separator is required")
    option_tokens = argv[:separator]
    command = argv[separator + 1 :]
    if not command or len(command) > 256:
        fail("a bounded renderer command is required")
    if any(not value or "\x00" in value or len(value) > 4096 for value in command):
        fail("renderer command arguments are invalid")
    if len(option_tokens) != 6:
        fail("exactly three resource-limit options are required")
    parsed: dict[str, str] = {}
    for index in range(0, len(option_tokens), 2):
        name, value = option_tokens[index : index + 2]
        if name not in {"--cpu-seconds", "--address-space-bytes", "--file-size-bytes"}:
            fail("an unknown resource-limit option was provided")
        if name in parsed:
            fail("resource-limit options must not be repeated")
        parsed[name] = value
    if len(parsed) != 3:
        fail("all resource-limit options are required")
    cpu_seconds = bounded_integer(parsed["--cpu-seconds"], "cpu-seconds", 1, MAX_CPU_SECONDS)
    address_space_bytes = bounded_integer(
        parsed["--address-space-bytes"],
        "address-space-bytes",
        MIN_ADDRESS_SPACE_BYTES,
        MAX_ADDRESS_SPACE_BYTES,
    )
    file_size_bytes = bounded_integer(
        parsed["--file-size-bytes"],
        "file-size-bytes",
        1,
        MAX_FILE_SIZE_BYTES,
    )
    return cpu_seconds, address_space_bytes, file_size_bytes, command


def set_limit(kind: int, value: int) -> None:
    current_soft, current_hard = resource.getrlimit(kind)
    del current_soft
    if current_hard != resource.RLIM_INFINITY and value > current_hard:
        fail("the host hard resource limit is below the required renderer limit", 70)
    resource.setrlimit(kind, (value, value))


def main() -> None:
    cpu_seconds, address_space_bytes, file_size_bytes, command = parse_arguments(sys.argv[1:])
    for required_name in ("RLIMIT_CPU", "RLIMIT_AS", "RLIMIT_FSIZE"):
        if not hasattr(resource, required_name):
            fail("the host does not provide required POSIX resource limits", 70)
    try:
        set_limit(resource.RLIMIT_CPU, cpu_seconds)
        set_limit(resource.RLIMIT_AS, address_space_bytes)
        set_limit(resource.RLIMIT_FSIZE, file_size_bytes)
        if hasattr(resource, "RLIMIT_CORE"):
            resource.setrlimit(resource.RLIMIT_CORE, (0, 0))
    except (OSError, ValueError):
        fail("the renderer resource limits could not be applied", 70)

    os.environ["PYTHONDONTWRITEBYTECODE"] = "1"
    try:
        os.execvp(command[0], command)
    except OSError:
        fail("the renderer process could not be executed", 70)


if __name__ == "__main__":
    main()
