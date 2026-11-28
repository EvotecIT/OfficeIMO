#!/usr/bin/env python3
"""Produce the small archive and comic fixtures with RARLAB's command-line programs."""

import argparse
import hashlib
import json
import os
import platform
import re
from pathlib import Path
import shutil
import struct
import subprocess
import tempfile
import zipfile
import zlib

ROOT = Path(__file__).resolve().parent.parent
ARCHIVES = ROOT / "OfficeIMO.Archives.Tests" / "Fixtures" / "IndependentRar5"
COMICS = ROOT / "OfficeIMO.Comic.Tests" / "Fixtures"


def rar_version(command):
    result = subprocess.run([*command, "-?"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, encoding="utf-8",
                            errors="replace", check=False)
    match = re.search(r"(?m)^RAR ([\d.]+(?: beta \d+)?)", result.stdout)
    if not match:
        raise RuntimeError(f"Could not identify RARLAB's rar tool at {command[-1]}")
    return match[1]


def generate(rar, legacy_rar):
    manifests = {ARCHIVES: [], COMICS: []}
    with tempfile.TemporaryDirectory(prefix="officeimo-rar-fixtures-") as scratch:
        work = Path(scratch)
        outputs = work / "out"
        outputs.mkdir()
        versions = {rar: rar_version([rar]), legacy_rar: rar_version([legacy_rar])}

        def run(command):
            subprocess.run(command, cwd=work, check=True)

        def add(name, switches, files, destination=ARCHIVES, tool=rar, password=None):
            command = ["a", "-cfg-", "-idq", *switches, f"out/{name}", *files]
            run([tool, *command])
            paths = sorted(outputs.glob(f"{Path(name).stem}.part*.rar")) if any(
                switch.startswith("-v") for switch in switches) else [outputs / name]
            if not paths:
                raise RuntimeError(f"rar produced no volumes for {name}")
            run([tool, "t", "-cfg-", "-idq", f"-p{password}" if password else "-p-", paths[0]])
            for path in paths:
                manifests[destination].append({
                    "file": path.name, "rarVersion": versions[tool], "arguments": command,
                    "sha256": hashlib.sha256(path.read_bytes()).hexdigest()
                })

        notes = work / "notes"
        (notes / "empty-dir").mkdir(parents=True)
        (notes / "café.txt").write_bytes(b"The north quay.\n" * 80)
        (notes / "señal.txt").write_bytes(b"The lighthouse is open.\n")
        (notes / "empty.txt").touch()
        (notes / "shortcut").symlink_to("café.txt")
        (work / "fixture.sfx").write_bytes(b"Inert OfficeIMO fixture prefix\n")
        (work / "zero.sfx").write_bytes(bytes(1024))
        (work / "comment.txt").write_bytes("RAR comment café\n".encode())
        piece = work / "piece.bin"
        piece.write_bytes(bytes(range(251)) * 7)
        shutil.copyfile(piece, work / "copy.bin")
        os.link(piece, work / "linked.bin")
        (work / "empty-directory").mkdir()
        (work / "directory-link").symlink_to("empty-directory", target_is_directory=True)
        add("directories.rar", ["-ma5", "-m0"], ["empty-directory", "notes/señal.txt"])
        add("stored.rar", ["-ma5", "-m0", "-qo+", "-ol"], ["notes"])
        add("plain.rar", ["-ma5", "-m0", "-qo-", "-tsm-"], ["notes/señal.txt"])
        # The tool adds the archive to the prefix. The prefix is inert and is never executed.
        add("prefixed.sfx", ["-ma5", "-m0", "-sfxfixture.sfx"], ["notes/señal.txt"])
        add("zero-prefixed.sfx", ["-ma5", "-m0", "-sfxzero.sfx"], ["notes/señal.txt"])
        add("legacy-prefixed.sfx", ["-ma4", "-m0", "-sfxfixture.sfx"], ["notes/señal.txt"], tool=legacy_rar)
        add("legacy.rar", ["-ma4", "-m0"], ["notes/señal.txt"], tool=legacy_rar)
        add("encrypted-headers.rar", ["-ma5", "-m0", "-hpfixture"], ["notes/café.txt"], password="fixture")
        add("encrypted-member.rar", ["-ma5", "-m0", "-pfixture"], ["notes/café.txt"], password="fixture")
        for method in range(1, 6):
            add(f"compressed-{method}.rar", ["-ma5", f"-m{method}"], ["notes/café.txt"])
        os.utime(piece, ns=(2200000001123456700, 2200000002123456700))
        add("metadata.rar", ["-ma5", "-m0", "-htb", "-qo+", "-tsmcap", "-ow", "-ol", "-oh", "-oi1:1", "-zcomment.txt"],
            ["piece.bin", "linked.bin", "copy.bin", "empty-directory", "directory-link"])
        add("split-blake.rar", ["-ma5", "-m0", "-v900b", "-htb"], ["piece.bin"])
        add("split-crc.rar", ["-ma5", "-m0", "-v900b", "-htc"], ["piece.bin"])
        with zipfile.ZipFile(COMICS / "harbour.cbz") as source:
            comic_files = ["ComicInfo.xml", "cover.png", "page10.png", "page2.png"]
            for name in comic_files:
                (work / name).write_bytes(source.read(name))
        add("harbour.cbr", ["-ma5", "-m0"], comic_files, destination=COMICS)
        add("harbour-compressed.cbr", ["-ma5", "-m3"], comic_files, destination=COMICS)

        for destination, manifest in manifests.items():
            for entry in manifest:
                shutil.copyfile(outputs / entry["file"], destination / entry["file"])
            evidence = {
                "producer": "RARLAB RAR command-line tool",
                "download": "https://www.rarlab.com/download.htm",
                "producerPlatform": f"{platform.system()} {platform.machine()}",
                "verification": "rar t passed for all archives and both complete volume sets.",
                "files": manifest
            }
            (destination / "rar5-provenance.json").write_text(json.dumps(evidence, indent=2, ensure_ascii=False) + "\n",
                                                              encoding="utf-8")
    generate_compressed_comment(rar)


def read_vint(data, position):
    value = 0
    for shift in range(0, 70, 7):
        byte = data[position]
        position += 1
        value |= (byte & 127) << shift
        if byte < 128:
            return value, position
    raise ValueError("Invalid RAR5 vint in producer output")


def find_rar5_block(data, block_type):
    if not data.startswith(b"Rar!\x1a\x07\x01\x00"):
        raise ValueError("Expected RAR5 output")
    start = 8
    while start < len(data):
        size, body = read_vint(data, start + 4)
        kind, position = read_vint(data, body)
        flags, position = read_vint(data, position)
        if flags & 1:
            _, position = read_vint(data, position)
        packed = read_vint(data, position)[0] if flags & 2 else 0
        header_end = body + size
        end = header_end + packed
        if kind == block_type:
            return start, body, header_end, end
        start = end
    raise ValueError(f"RAR5 producer output has no block of type {block_type}")


def generate_compressed_comment(rar):
    # Modern rar stores -z comments. Relocate a genuine method-3 file record named CMT
    # to the comment position, change its type to service, and refresh the header CRC.
    with tempfile.TemporaryDirectory(prefix="officeimo-rar-comment-") as scratch:
        work = Path(scratch)
        comment = b"Compressible RAR comment.\n" * 300
        (work / "comment.txt").write_bytes(comment)
        (work / "CMT").write_bytes(comment)
        (work / "page.txt").write_bytes(b"Member payload.\n" * 100)
        base_args = ["a", "-cfg-", "-idq", "-ma5", "-m0", "-qo-", "-tsm-", "-zcomment.txt", "base.rar", "page.txt"]
        donor_args = ["a", "-cfg-", "-idq", "-ma5", "-m3", "-md128k", "-qo-", "-tsm-", "-zcomment.txt", "donor.rar", "CMT"]
        for command in (base_args, donor_args):
            subprocess.run([rar, *command], cwd=work, check=True)
        base = (work / "base.rar").read_bytes()
        donor = (work / "donor.rar").read_bytes()
        start, body, header_end, end = find_rar5_block(donor, 2)
        if end - header_end >= len(comment):
            raise ValueError("rar did not compress the comment payload")
        record = bytearray(donor[start + 4:end])
        record[body - start - 4] = 3
        crc = struct.pack("<I", zlib.crc32(record[:header_end - start - 4]))
        base_start, _, _, base_end = find_rar5_block(base, 3)
        result = base[:base_start] + crc + record + base[base_end:]
        name = "compressed-comment.rar"
        (work / name).write_bytes(result)
        subprocess.run([rar, "t", "-cfg-", "-idq", name], cwd=work, check=True)
        subprocess.run([rar, "cw", "-cfg-", "-idq", name, "decoded.txt"], cwd=work, check=True,
                       stdout=subprocess.DEVNULL)
        if (work / "decoded.txt").read_bytes() != comment:
            raise ValueError("rar did not recover the original comment")
        (ARCHIVES / name).write_bytes(result)
        evidence = {
            "producer": "RARLAB RAR command-line tool",
            "producerPlatform": f"{platform.system()} {platform.machine()}",
            "rarVersion": rar_version([rar]), "file": name, "arguments": [base_args, donor_args],
            "derivation": "Replaced the stored comment with the producer's compressed CMT file record, changing its type to service and updating the header CRC.",
            "verification": "rar t passed and rar cw recovered the original comment byte-for-byte.",
            "sha256": hashlib.sha256(result).hexdigest()
        }
        (ARCHIVES / "compressed-comment-provenance.json").write_text(json.dumps(evidence, indent=2) + "\n", encoding="utf-8")


def generate_windows_links(rar, wine):
    if os.name != "nt" and not wine:
        raise RuntimeError("Run on Windows or provide a Wine 11 executable with reparse point support")
    launcher = [wine] if wine else []
    with tempfile.TemporaryDirectory(prefix="officeimo-rar-windows-") as scratch:
        work = Path(scratch)
        target = work / "directory"
        target.mkdir()
        (target / "page.txt").write_bytes(b"Windows link target.\n")
        windows_target = "Z:" + str(target).replace("/", "\\") if wine else str(target)
        cmd = launcher + ["cmd.exe", "/d", "/c"]
        subprocess.run(cmd + ["mklink", "file-link", "directory\\page.txt"], cwd=work, check=True)
        subprocess.run(cmd + ["mklink", "/J", "junction", windows_target], cwd=work, check=True)
        try:
            args = ["a", "-cfg-", "-idq", "-ma5", "-m0", "-ola", "windows-links.rar", "directory", "file-link", "junction"]
            subprocess.run(launcher + [rar, *args], cwd=work, check=True)
            archive = work / "windows-links.rar"
            command = launcher + [rar, "lt", "-cfg-", archive.name]
            output = subprocess.check_output(command, cwd=work, encoding="utf-8", errors="replace")
            if "windows symbolic link" not in output.lower() or "junction" not in output.lower():
                raise RuntimeError("RAR did not record both Windows redirection types:\n" + output)
            subprocess.run(launcher + [rar, "t", "-cfg-", "-idq", archive.name], cwd=work, check=True)
            shutil.copyfile(archive, ARCHIVES / archive.name)
            environment = subprocess.check_output([wine, "--version"], text=True).strip() if wine else "Windows NTFS"
            evidence = {
                "producer": "RARLAB RAR command-line tool for Windows",
                "producerPlatform": f"{platform.machine()}, {environment}",
                "rarVersion": rar_version(launcher + [rar]),
                "file": archive.name, "arguments": args, "verification": "RAR listed both link types and rar t passed.",
                "sha256": hashlib.sha256(archive.read_bytes()).hexdigest()
            }
            (ARCHIVES / "windows-links-provenance.json").write_text(json.dumps(evidence, indent=2) + "\n", encoding="utf-8")
        finally:
            subprocess.run(cmd + ["rmdir", "junction"], cwd=work, check=True)


def generate_numeric_owner(rar, qemu):
    if platform.system() != "Linux" or os.geteuid() != 0:
        raise RuntimeError("Run the numeric owner fixture in an isolated Linux guest with permission to chown")
    launcher = [qemu] if qemu else []
    with tempfile.TemporaryDirectory(prefix="officeimo-rar-owner-") as scratch:
        work = Path(scratch)
        item = work / "owner.txt"
        item.write_bytes(b"Numeric owner.\n")
        os.chown(item, 42421, 42422)
        args = ["a", "-cfg-", "-idq", "-ma5", "-m0", "-ow", "numeric-owner.rar", item.name]
        archive = work / "numeric-owner.rar"
        subprocess.run(launcher + [rar, *args], cwd=work, check=True)
        output = subprocess.check_output(launcher + [rar, "lt", "-cfg-", archive.name],
                                         cwd=work, encoding="utf-8", errors="replace")
        if "#42421:#42422" not in output:
            raise RuntimeError("RAR did not preserve the numerical Unix IDs:\n" + output)
        subprocess.run(launcher + [rar, "t", "-cfg-", "-idq", archive.name], cwd=work, check=True)
        shutil.copyfile(archive, ARCHIVES / archive.name)
        evidence = {
            "producer": "RARLAB RAR command-line tool for Linux",
            "producerPlatform": f"{platform.machine()}, Linux" + (" with QEMU user-mode execution" if qemu else ""),
            "rarVersion": rar_version(launcher + [rar]), "file": archive.name, "arguments": args,
            "verification": "RAR listed the numeric Unix IDs and rar t passed.",
            "sha256": hashlib.sha256(archive.read_bytes()).hexdigest()
        }
        (ARCHIVES / "numeric-owner-provenance.json").write_text(json.dumps(evidence, indent=2) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rar", required=True, type=Path, help="RARLAB rar executable for RAR5")
    parser.add_argument("--rar4", type=Path, help="RARLAB rar 6.x executable for producing RAR4")
    parser.add_argument("--windows-links", action="store_true", help="Produce the NTFS redirection fixture only")
    parser.add_argument("--numeric-owner", action="store_true", help="Produce the Unix numeric owner fixture only")
    parser.add_argument("--compressed-comment", action="store_true", help="Produce the compressed RAR5 comment fixture only")
    parser.add_argument("--qemu", type=Path, help="Optional QEMU executable for the Linux rar binary")
    parser.add_argument("--wine", type=Path, help="Wine executable for the Windows redirection fixture")
    args = parser.parse_args()
    if args.windows_links:
        generate_windows_links(args.rar.resolve(), args.wine.resolve() if args.wine else None)
    elif args.numeric_owner:
        generate_numeric_owner(args.rar.resolve(), args.qemu.resolve() if args.qemu else None)
    elif args.compressed_comment:
        generate_compressed_comment(args.rar.resolve())
    elif args.rar4:
        generate(args.rar.resolve(), args.rar4.resolve())
    else:
        parser.error("--rar4 is required when producing the main fixture set")


if __name__ == "__main__":
    main()
