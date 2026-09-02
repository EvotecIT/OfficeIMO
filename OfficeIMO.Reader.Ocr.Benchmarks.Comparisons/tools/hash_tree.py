#!/usr/bin/env python3
"""Emit a deterministic SHA-256 provenance digest for one extracted payload tree."""

from __future__ import annotations

import argparse
import hashlib
import os
from pathlib import Path


def file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def tree_digest(root: Path) -> tuple[str, int]:
    root = root.resolve()
    records: list[bytes] = []
    for directory, names, files in os.walk(root, followlinks=False):
        names.sort()
        files.sort()
        for name in [*names, *files]:
            path = Path(directory) / name
            relative = path.relative_to(root).as_posix()
            if path.is_symlink():
                record = f"L\0{relative}\0{os.readlink(path)}"
            elif path.is_file():
                record = f"F\0{relative}\0{file_digest(path)}"
            elif path.is_dir():
                continue
            else:
                raise RuntimeError(f"Unsupported payload entry: {path}")
            records.append(record.encode("utf-8"))

    digest = hashlib.sha256(b"\n".join(records)).hexdigest()
    return digest, len(records)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("root", type=Path)
    args = parser.parse_args()
    if not args.root.is_dir():
        raise NotADirectoryError(args.root)
    digest, count = tree_digest(args.root)
    print(f"{digest} {count}")


if __name__ == "__main__":
    main()
