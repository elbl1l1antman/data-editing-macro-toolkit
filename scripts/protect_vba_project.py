#!/usr/bin/env python3
"""Password-protect an Office VBA project without opening the VBE UI."""

from __future__ import annotations

import argparse
import io
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path


VENDOR_WHEEL = Path(__file__).with_name("vendor") / "olefile-0.47-py2.py3-none-any.whl"
if VENDOR_WHEEL.exists():
    sys.path.insert(0, str(VENDOR_WHEEL))

import olefile  # type: ignore  # noqa: E402


def encrypt_project_value(data: bytes, project_id: bytes, seed: int) -> bytes:
    """Encode a PROJECT value using MS-OVBA 2.4.3.2."""
    project_key = sum(project_id) & 0xFF
    version_enc = seed ^ 2
    project_key_enc = seed ^ project_key
    output = bytearray((seed, version_enc, project_key_enc))
    unencrypted_1 = project_key
    encrypted_1 = project_key_enc
    encrypted_2 = version_enc

    def append_encrypted(value: int) -> None:
        nonlocal unencrypted_1, encrypted_1, encrypted_2
        encoded = value ^ ((encrypted_2 + unencrypted_1) & 0xFF)
        output.append(encoded)
        encrypted_2, encrypted_1, unencrypted_1 = encrypted_1, encoded, value

    for _ in range((seed & 6) // 2):
        append_encrypted(0)
    for value in len(data).to_bytes(4, "little"):
        append_encrypted(value)
    for value in data:
        append_encrypted(value)
    return output.hex().upper().encode("ascii")


def protected_project_stream(project: bytes, password: str) -> bytes:
    project_id_match = re.search(br'^ID="([^"]+)"', project, re.MULTILINE)
    if not project_id_match:
        raise ValueError("PROJECT stream has no project ID")
    project_id = project_id_match.group(1)

    replacements = {
        br'CMG="[0-9A-F]+"': b'CMG="' + encrypt_project_value(b"\x05\x00\x00\x00", project_id, 0x06) + b'"',
        br'DPB="[0-9A-F]+"': b'DPB="' + encrypt_project_value(password.encode("cp949") + b"\x00", project_id, 0x08) + b'"',
        br'GC="[0-9A-F]+"': b'GC="' + encrypt_project_value(b"\x00", project_id, 0x08) + b'"',
    }

    updated = project
    for pattern, replacement in replacements.items():
        updated, count = re.subn(pattern, lambda _match, value=replacement: value, updated, count=1)
        if count != 1:
            raise ValueError(f"PROJECT stream is missing {pattern!r}")

    growth = len(updated) - len(project)
    if growth > 0:
        workspace_at = updated.find(b"[Workspace]")
        if workspace_at < 0:
            raise ValueError("PROJECT stream has no workspace section for fixed-size patching")
        prefix = updated[:workspace_at]
        workspace = updated[workspace_at:]
        available = workspace.count(b", ")
        if available < growth:
            raise ValueError("PROJECT stream has insufficient optional whitespace for protection data")
        workspace = workspace.replace(b", ", b",", growth)
        updated = prefix + workspace

    if len(updated) != len(project):
        raise ValueError("Protected PROJECT stream size changed")
    return updated


def protect_package(path: Path, password: str) -> None:
    with zipfile.ZipFile(path, "r") as archive:
        entries = [(item, archive.read(item.filename)) for item in archive.infolist()]

    replaced = False
    rewritten: list[tuple[zipfile.ZipInfo, bytes]] = []
    for item, data in entries:
        if item.filename == "xl/vbaProject.bin":
            compound = io.BytesIO(data)
            with olefile.OleFileIO(compound, write_mode=True) as ole:
                project = ole.openstream("PROJECT").read()
                ole.write_stream("PROJECT", protected_project_stream(project, password))
            data = compound.getvalue()
            replaced = True
        rewritten.append((item, data))
    if not replaced:
        raise ValueError(f"No xl/vbaProject.bin found in {path}")

    with tempfile.NamedTemporaryFile(delete=False, dir=path.parent, suffix=path.suffix) as temp:
        temp_path = Path(temp.name)
    try:
        with zipfile.ZipFile(temp_path, "w") as archive:
            for item, data in rewritten:
                archive.writestr(item, data)
        os.replace(temp_path, path)
    finally:
        temp_path.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("package", type=Path)
    parser.add_argument("--password-env", default="DATA_EDITING_TOOLKIT_VBA_PASSWORD")
    args = parser.parse_args()
    password = os.environ.get(args.password_env, "")
    if not password:
        raise SystemExit(f"Missing environment variable: {args.password_env}")
    protect_package(args.package, password)
    print(f"Protected VBA project: {args.package}")


if __name__ == "__main__":
    main()
