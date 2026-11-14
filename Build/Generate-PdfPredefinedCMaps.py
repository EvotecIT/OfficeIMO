"""Generate OfficeIMO.Pdf predefined CJK CMap data from Adobe's published CMap and CID-to-Unicode resources.

Each supported Adobe character collection becomes one embedded resource holding a small index and one
raw-deflate segment per predefined CMap, plus the collection's CID-to-Unicode map. The reader inflates
only the segment a document names, the first time it is needed. Segments store our own run encoding:
code-to-CID runs grouped by code length, notdef runs, the codespace ranges and an optional usecmap
parent, so vertical CMaps keep only their overrides.

Usage (Python 3, no third-party packages):
    python Build/Generate-PdfPredefinedCMaps.py <cmap-resources checkout> <mapping-resources-pdf checkout>

Inputs are https://github.com/adobe-type-tools/cmap-resources and
https://github.com/adobe-type-tools/mapping-resources-pdf (both BSD-3-Clause). The script writes
OfficeIMO.Pdf/Resources/AdobeCMaps/Adobe-<Ordering>.bin, source.json with the upstream revisions and
input hashes, and OfficeIMO.Pdf/Reading/Font/PdfPredefinedCMap.Catalog.Generated.cs.
"""
import hashlib
import json
import re
import subprocess
import sys
import zlib
from pathlib import Path

# Character collections whose CMaps and Unicode maps are both published. Identity and Manga1 are excluded:
# Identity-H/V are handled without data and Manga1 has no published CID-to-Unicode map.
COLLECTIONS = [
    ("GB1", "Adobe-GB1-6"),
    ("CNS1", "Adobe-CNS1-7"),
    ("Japan1", "Adobe-Japan1-7"),
    ("Korea1", "Adobe-Korea1-2"),
    ("KR", "Adobe-KR-9"),
]
MAGIC = b"OCMP"
FORMAT_VERSION = 1
UNICODE_ENTRY = "#Unicode"
COPYRIGHT = ("Mapping data derived from Adobe CMap Resources and Adobe Mapping Resources for PDF. "
             "Copyright 1990-2023 Adobe. All rights reserved. Distributed under the BSD-3-Clause license.")

TOKEN = re.compile(rb"<[0-9A-Fa-f\s]*>|/[^\s/<>\[\]()]+|\[|\]|[^\s/<>\[\]()%]+|%[^\r\n]*")
REPO = Path(__file__).resolve().parent.parent
OUTPUT = REPO / "OfficeIMO.Pdf" / "Resources" / "AdobeCMaps"
CATALOG = REPO / "OfficeIMO.Pdf" / "Reading" / "Font" / "PdfPredefinedCMap.Catalog.Generated.cs"


def tokens(data):
    for match in TOKEN.finditer(data):
        token = match.group(0)
        if not token.startswith(b"%"):
            yield token.decode("ascii")


def hex_code(token):
    digits = re.sub(r"\s", "", token[1:-1])
    if not digits or len(digits) % 2:
        raise ValueError(f"Invalid hex token {token}")
    return len(digits) // 2, int(digits, 16)


class CMap:
    def __init__(self, name):
        self.name = name
        self.wmode = 0
        self.parent = None
        self.codespace = []
        self.cids = {}
        self.notdef = {}


def parse_cmap(path):
    cmap = CMap(path.name)
    items = list(tokens(path.read_bytes()))
    index = 0
    previous = []
    while index < len(items):
        token = items[index]
        if token == "usecmap":
            cmap.parent = previous[-1][1:]
        elif token == "/WMode":
            cmap.wmode = int(items[index + 1])
        elif token in ("begincodespacerange", "begincidrange", "begincidchar", "beginnotdefrange"):
            end = "end" + token[5:]
            index += 1
            body = []
            while items[index] != end:
                body.append(items[index])
                index += 1
            if token == "begincodespacerange":
                for low, high in zip(body[0::2], body[1::2]):
                    low_length, low_value = hex_code(low)
                    high_length, high_value = hex_code(high)
                    if low_length != high_length:
                        raise ValueError(f"{path}: codespace length mismatch")
                    cmap.codespace.append((low_length, low_value, high_value))
            elif token == "begincidchar":
                for code, cid in zip(body[0::2], body[1::2]):
                    cmap.cids[hex_code(code)] = int(cid)
            else:
                target = cmap.cids if token == "begincidrange" else cmap.notdef
                for low, high, cid in zip(body[0::3], body[1::3], body[2::3]):
                    length, first = hex_code(low)
                    high_length, last = hex_code(high)
                    if length != high_length or last < first:
                        raise ValueError(f"{path}: invalid range {low} {high}")
                    for offset in range(last - first + 1):
                        target[(length, first + offset)] = int(cid) + (offset if token == "begincidrange" else 0)
        previous.append(token)
        index += 1
    return cmap


def parse_unicode(path):
    values = {}
    items = list(tokens(path.read_bytes()))
    index = 0
    while index < len(items):
        token = items[index]
        if token in ("beginbfchar", "beginbfrange"):
            end = "end" + token[5:]
            index += 1
            body = []
            while items[index] != end:
                body.append(items[index])
                index += 1
            if token == "beginbfchar":
                for code, value in zip(body[0::2], body[1::2]):
                    values[hex_code(code)[1]] = utf16_scalars(value)
            else:
                position = 0
                while position < len(body):
                    first = hex_code(body[position])[1]
                    last = hex_code(body[position + 1])[1]
                    if body[position + 2] == "[":
                        position += 3
                        cid = first
                        while body[position] != "]":
                            values[cid] = utf16_scalars(body[position])
                            cid += 1
                            position += 1
                        position += 1
                    else:
                        start = utf16_scalars(body[position + 2])
                        for offset in range(last - first + 1):
                            values[first + offset] = start[:-1] + [start[-1] + offset]
                        position += 3
        index += 1
    return values


def utf16_scalars(token):
    digits = re.sub(r"\s", "", token[1:-1])
    units = [int(digits[i:i + 4], 16) for i in range(0, len(digits), 4)]
    text = b"".join(unit.to_bytes(2, "big") for unit in units).decode("utf-16-be")
    return [ord(character) for character in text]


def varint(value, out):
    if value < 0:
        raise ValueError("negative varint")
    while True:
        byte = value & 0x7F
        value >>= 7
        if value:
            out.append(byte | 0x80)
        else:
            out.append(byte)
            return


def zigzag(value, out):
    varint((value << 1) if value >= 0 else ((-value << 1) - 1), out)


def text(value, out):
    data = value.encode("ascii")
    varint(len(data), out)
    out.extend(data)


def runs(mapping, constant):
    """Groups (length, code) -> cid entries into maximal runs of consecutive codes per code length."""
    groups = {}
    for (length, code) in sorted(mapping):
        cid = mapping[(length, code)]
        group = groups.setdefault(length, [])
        if group:
            first, last, first_cid = group[-1]
            expected = first_cid if constant else first_cid + (last - first) + 1
            if code == last + 1 and cid == expected:
                group[-1] = (first, code, first_cid)
                continue
        group.append((code, code, cid))
    return groups


def encode_runs(mapping, constant, out):
    groups = runs(mapping, constant)
    varint(len(groups), out)
    for length in sorted(groups):
        out.append(length)
        group = groups[length]
        varint(len(group), out)
        next_code = 0
        next_cid = 0
        for first, last, cid in group:
            varint(first - next_code, out)
            varint(last - first, out)
            zigzag(cid - next_cid, out)
            next_code = last + 1
            next_cid = cid if constant else cid + (last - first) + 1


def encode_cmap(cmap):
    out = bytearray()
    varint(cmap.wmode, out)
    text(cmap.parent or "", out)
    varint(len(cmap.codespace), out)
    for length, low, high in cmap.codespace:
        out.append(length)
        out.extend(low.to_bytes(length, "big"))
        out.extend(high.to_bytes(length, "big"))
    encode_runs(cmap.cids, False, out)
    encode_runs(cmap.notdef, True, out)
    return bytes(out)


def encode_unicode(values):
    out = bytearray()
    varint(max(values) + 1, out)
    entries = []
    for cid in sorted(values):
        scalars = values[cid]
        if len(scalars) == 1 and entries and entries[-1][0] == "run":
            _, first, count, scalar = entries[-1]
            if cid == first + count and scalars[0] == scalar + count:
                entries[-1] = ("run", first, count + 1, scalar)
                continue
        entries.append(("run", cid, 1, scalars[0]) if len(scalars) == 1 else ("sequence", cid, scalars))
    varint(len(entries), out)
    next_cid = 0
    previous_scalar = 0
    for entry in entries:
        varint(entry[1] - next_cid, out)
        if entry[0] == "run":
            _, cid, count, scalar = entry
            varint((count - 1) << 1, out)
            zigzag(scalar - previous_scalar, out)
            next_cid = cid + count
            previous_scalar = scalar + count - 1
        else:
            _, cid, scalars = entry
            varint((len(scalars) << 1) | 1, out)
            for scalar in scalars:
                zigzag(scalar - previous_scalar, out)
                previous_scalar = scalar
            next_cid = cid + 1
    return bytes(out)


class Reader:
    def __init__(self, data):
        self.data = data
        self.position = 0

    def byte(self):
        value = self.data[self.position]
        self.position += 1
        return value

    def varint(self):
        value = shift = 0
        while True:
            byte = self.byte()
            value |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                return value

    def zigzag(self):
        value = self.varint()
        return (value >> 1) if not value & 1 else -((value + 1) >> 1)


def decode_runs(reader, constant):
    mapping = {}
    for _ in range(reader.varint()):
        length = reader.byte()
        next_code = next_cid = 0
        for _ in range(reader.varint()):
            first = next_code + reader.varint()
            last = first + reader.varint()
            cid = next_cid + reader.zigzag()
            for offset in range(last - first + 1):
                mapping[(length, first + offset)] = cid if constant else cid + offset
            next_code = last + 1
            next_cid = cid if constant else cid + (last - first) + 1
    return mapping


def verify_cmap(cmap, payload):
    reader = Reader(payload)
    assert reader.varint() == cmap.wmode
    parent_length = reader.varint()
    parent = reader.data[reader.position:reader.position + parent_length].decode("ascii")
    reader.position += parent_length
    assert parent == (cmap.parent or "")
    codespace = []
    for _ in range(reader.varint()):
        length = reader.byte()
        low = int.from_bytes(reader.data[reader.position:reader.position + length], "big")
        high = int.from_bytes(reader.data[reader.position + length:reader.position + 2 * length], "big")
        reader.position += 2 * length
        codespace.append((length, low, high))
    assert codespace == cmap.codespace
    assert decode_runs(reader, False) == cmap.cids, cmap.name
    assert decode_runs(reader, True) == cmap.notdef, cmap.name
    assert reader.position == len(payload)


def verify_unicode(values, payload):
    reader = Reader(payload)
    reader.varint()
    decoded = {}
    next_cid = previous_scalar = 0
    for _ in range(reader.varint()):
        cid = next_cid + reader.varint()
        header = reader.varint()
        if header & 1:
            scalars = []
            for _ in range(header >> 1):
                previous_scalar += reader.zigzag()
                scalars.append(previous_scalar)
            decoded[cid] = scalars
            next_cid = cid + 1
        else:
            count = (header >> 1) + 1
            scalar = previous_scalar + reader.zigzag()
            for offset in range(count):
                decoded[cid + offset] = [scalar + offset]
            next_cid = cid + count
            previous_scalar = scalar + count - 1
    assert decoded == values
    assert reader.position == len(payload)


def deflate(data):
    compressor = zlib.compressobj(9, zlib.DEFLATED, -15, 9)
    return compressor.compress(data) + compressor.flush()


def container(entries):
    index = bytearray()
    varint(FORMAT_VERSION, index)
    copyright_bytes = COPYRIGHT.encode("ascii")
    varint(len(copyright_bytes), index)
    index.extend(copyright_bytes)
    varint(len(entries), index)
    offset = 0
    compressed = []
    for name, payload in entries:
        packed = deflate(payload)
        text(name, index)
        varint(offset, index)
        varint(len(packed), index)
        varint(len(payload), index)
        compressed.append(packed)
        offset += len(packed)
    header = bytearray(MAGIC)
    varint(len(index), header)
    return bytes(header + index + b"".join(compressed))


def revision(path):
    return subprocess.run(["git", "-C", str(path), "rev-parse", "HEAD"], check=True,
                          capture_output=True, text=True).stdout.strip()


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    cmap_root = Path(sys.argv[1])
    unicode_root = Path(sys.argv[2])
    cmap_revision = revision(cmap_root)
    unicode_revision = revision(unicode_root)
    catalog = []
    outputs = []
    inputs = []
    for ordering, directory in COLLECTIONS:
        files = sorted((cmap_root / directory / "CMap").iterdir(), key=lambda item: item.name)
        cmaps = [parse_cmap(path) for path in files]
        names = {cmap.name for cmap in cmaps}
        entries = []
        for cmap, path in zip(cmaps, files):
            if cmap.parent and cmap.parent not in names:
                raise ValueError(f"{cmap.name} uses {cmap.parent} outside {directory}")
            if not cmap.codespace and not cmap.parent:
                raise ValueError(f"{cmap.name} has no codespace")
            payload = encode_cmap(cmap)
            verify_cmap(cmap, payload)
            entries.append((cmap.name, payload))
            catalog.append((cmap.name, ordering))
            inputs.append({"file": f"{directory}/CMap/{path.name}", "sha256": sha256(path)})
        unicode_path = unicode_root / "pdf2unicode" / f"Adobe-{ordering}-UCS2"
        values = parse_unicode(unicode_path)
        unicode_payload = encode_unicode(values)
        verify_unicode(values, unicode_payload)
        entries.append((UNICODE_ENTRY, unicode_payload))
        inputs.append({"file": f"pdf2unicode/{unicode_path.name}", "sha256": sha256(unicode_path)})
        resource = OUTPUT / f"Adobe-{ordering}.bin"
        resource.write_bytes(container(entries))
        outputs.append({"file": resource.name, "collection": f"Adobe-{ordering}", "cmaps": len(cmaps),
                        "sha256": sha256(resource), "size": resource.stat().st_size})
        print(f"{resource.name}: {len(cmaps)} CMaps, {resource.stat().st_size} bytes")

    for license_file, source in (("LICENSE.md", cmap_root / "LICENSE.md"), ("LICENSE.txt", unicode_root / "LICENSE.txt")):
        (OUTPUT / license_file).write_bytes(source.read_bytes())
    source = {
        "generator": "Build/Generate-PdfPredefinedCMaps.py",
        "role": "Predefined CJK CMaps (code to CID) and Adobe CID-to-Unicode maps in OfficeIMO's compact run format",
        "upstream": {
            "cmap-resources": {"url": "https://github.com/adobe-type-tools/cmap-resources", "revision": cmap_revision,
                               "license": "BSD-3-Clause", "licenseFile": "LICENSE.md"},
            "mapping-resources-pdf": {"url": "https://github.com/adobe-type-tools/mapping-resources-pdf",
                                      "revision": unicode_revision, "license": "BSD-3-Clause", "licenseFile": "LICENSE.txt"},
        },
        "outputs": outputs,
        "inputs": inputs,
    }
    (OUTPUT / "source.json").write_text(json.dumps(source, indent=2) + "\n", encoding="utf-8", newline="\n")

    lines = [
        "// <auto-generated>",
        "// Generated by Build/Generate-PdfPredefinedCMaps.py from Adobe CMap Resources",
        f"// ({cmap_revision}) and Adobe Mapping Resources for PDF ({unicode_revision}).",
        "// The Adobe data is distributed under the BSD-3-Clause license; see THIRD-PARTY-NOTICES.md.",
        "// </auto-generated>",
        "",
        "#nullable enable",
        "",
        "namespace OfficeIMO.Pdf {",
        "",
        "    internal sealed partial class PdfPredefinedCMap {",
        "        /// <summary>Returns the Adobe character collection ordering of a bundled predefined CMap name.</summary>",
        "        private static string? GetBundledOrdering(string name) => name switch {",
    ]
    for name, ordering in sorted(catalog):
        lines.append(f"            \"{name}\" => \"{ordering}\",")
    lines += [
        "            _ => null",
        "        };",
        "",
        "        /// <summary>Orderings that have a bundled CID-to-Unicode map.</summary>",
        "        private static bool IsBundledOrdering(string ordering) => ordering is " +
        " or ".join(f"\"{ordering}\"" for ordering, _ in COLLECTIONS) + ";",
        "    }",
        "}",
        "",
    ]
    CATALOG.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
