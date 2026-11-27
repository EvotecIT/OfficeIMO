#!/usr/bin/env python3
"""Derives the damaged DBX and OLM recovery fixtures from the committed good fixtures.

Run from any directory: python3 OfficeIMO.Email.Tests/Corpora/damaged-stores/generate.py
Only the standard library is used. Every damage is a byte-level edit of a good fixture, so the
expected outcome of each file follows from the edit described in fixtures.json.
"""

import hashlib
import io
import json
import struct
import warnings
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
CORPORA = HERE.parent
DBX_SOURCE = CORPORA / "dbx-metadata-unknown-producer" / "Inbox.dbx"
OLM_SOURCE = CORPORA / "native-apple-20261003" / "outlook" / "native-outlook.olm"

# Inbox.dbx layout (see dbx-metadata-unknown-producer/provenance.json): one message record at 0x2e10
# whose body spans 20 linked fragments of 0x200 payload bytes, the first at 0xead4 and each 0x210 apart.
DBX_RECORD = 0x2E10
DBX_FIRST_FRAGMENT = 0xEAD4
DBX_FRAGMENT_STRIDE = 0x210
DBX_HEADER_ROOT = 0xE4


def fragment(index):
    return DBX_FIRST_FRAGMENT + index * DBX_FRAGMENT_STRIDE


def put_u32(data, offset, value):
    struct.pack_into("<I", data, offset, value)


def dbx_variants(source):
    variants = {}

    truncated_at = fragment(7) + 16 + 100
    variants["Inbox-truncated.dbx"] = (
        bytes(source[:truncated_at]),
        "Truncated 100 bytes into the payload of the eighth fragment; the active index at the end of the file is gone.",
    )

    broken = bytearray(source)
    put_u32(broken, fragment(4) + 12, 0x00012340)
    variants["Inbox-broken-link.dbx"] = (
        bytes(broken),
        "The fifth fragment's next pointer is redirected to 0x12340, which holds no fragment.",
    )

    no_index = bytearray(source)
    put_u32(no_index, DBX_HEADER_ROOT, 0x00ABCDE0)
    variants["Inbox-no-index.dbx"] = (
        bytes(no_index),
        "The header's active-index root pointer is redirected outside the file; records and fragments are intact.",
    )

    no_record = bytearray(source)
    put_u32(no_record, DBX_RECORD, 0)
    variants["Inbox-lost-record.dbx"] = (
        bytes(no_record),
        "The message record's self-address marker is cleared; the index still lists it and its fragments are intact.",
    )
    return variants


def local_entries(data):
    """Yields (offset, name, data_offset, compressed_size) for each file entry of a good archive.

    Outlook writes ZIP64 local headers, so sizes come from the central directory.
    """
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for info in archive.infolist():
            name_length, extra_length = struct.unpack_from("<HH", data, info.header_offset + 26)
            data_offset = info.header_offset + 30 + name_length + extra_length
            yield info.header_offset, info.filename, data_offset, info.compress_size


def central_directory_offset(data):
    """Finds the first central directory header after the last entry's data (ZIP64 end records hide the offset)."""
    end = max(data_offset + size for _, _, data_offset, size in local_entries(data))
    return data.index(b"PK\x01\x02", end)


def olm_variants(source):
    variants = {}
    entries = {name: (offset, data_offset, size) for offset, name, data_offset, size in local_entries(source)}
    message = "Local/com.microsoft.__Messages/OfficeIMO native Outlook validation PST/OfficeIMO Mail/message_00000.xml"
    attachment = ("Local/com.microsoft.__Messages/OfficeIMO native Outlook validation PST/OfficeIMO Mail/"
                  "com.microsoft.__Attachments/0E6A6136-1050-4252-91CC-3280DB8783C4@example.test_0000")
    contacts = "Local/Książka adresowa/Contacts.xml"

    variants["native-outlook-no-central-directory.olm"] = (
        bytes(source[:central_directory_offset(source)]),
        "The central directory and end record are removed; every local entry is intact.",
    )

    _, data_offset, _ = entries[message]
    variants["native-outlook-truncated.olm"] = (
        bytes(source[:data_offset + 2000]),
        "Truncated 2000 bytes into the only mail item's XML entry, which is the last file entry.",
    )

    bad_attachment = bytearray(source)
    _, data_offset, _ = entries[attachment]
    bad_attachment[data_offset] ^= 0x20
    variants["native-outlook-attachment-crc.olm"] = (
        bytes(bad_attachment),
        "One byte of the mail item's first attachment is changed, so its CRC-32 no longer matches.",
    )

    bad_contact = bytearray(source)
    _, data_offset, size = entries[contacts]
    text = bytes(source[data_offset:data_offset + size])
    marker = text.index(b"<OPFContactCopyDisplayName")
    value = text.index(b">", marker) + 1
    bad_contact[data_offset + value] ^= 0x20
    variants["native-outlook-contacts-crc.olm"] = (
        bytes(bad_contact),
        "The case of the first character of the contact's display name is flipped; the XML stays well formed "
        "but its CRC-32 no longer matches.",
    )

    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        output = io.BytesIO()

        class Unseekable(io.RawIOBase):
            def writable(self):
                return True

            def write(self, data):
                return output.write(data)

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with zipfile.ZipFile(Unseekable(), "w", compression=zipfile.ZIP_DEFLATED) as rewritten:
                for info in archive.infolist():
                    copy = zipfile.ZipInfo(info.filename, info.date_time)
                    copy.compress_type = zipfile.ZIP_STORED if info.is_dir() else zipfile.ZIP_DEFLATED
                    rewritten.writestr(copy, archive.read(info))
        deflated = output.getvalue()
    variants["native-outlook-deflate-no-central-directory.olm"] = (
        deflated[:central_directory_offset(deflated)],
        "Every file entry is recompressed with Deflate and sized only by data descriptors; the central directory "
        "and end record are then removed.",
    )
    return variants


def main():
    dbx = DBX_SOURCE.read_bytes()
    olm = OLM_SOURCE.read_bytes()
    manifest = {
        "schemaVersion": 1,
        "generator": "generate.py",
        "sources": {
            "dbx": {"path": "../dbx-metadata-unknown-producer/Inbox.dbx", "sha256": hashlib.sha256(dbx).hexdigest(),
                    "license": "MIT; see ../dbx-metadata-unknown-producer/LICENSE"},
            "olm": {"path": "../native-apple-20261003/outlook/native-outlook.olm",
                    "sha256": hashlib.sha256(olm).hexdigest()},
        },
        "files": [],
    }
    for name, (data, damage) in {**dbx_variants(dbx), **olm_variants(olm)}.items():
        (HERE / name).write_bytes(data)
        manifest["files"].append({"path": name, "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest(),
                                  "damage": damage})
    (HERE / "fixtures.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
