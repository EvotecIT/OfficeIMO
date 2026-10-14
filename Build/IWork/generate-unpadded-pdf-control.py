"""Generate the synthetic one-cell Numbers PDF control; no Apple producer is claimed."""
from pathlib import Path
import argparse
import struct
import zipfile


def varint(value):
    data = bytearray()
    while value > 127:
        data.append((value & 127) | 128)
        value >>= 7
    return bytes(data + bytes([value]))


def integer(field, value):
    return varint(field << 3) + varint(value)


def blob(field, value):
    return varint((field << 3) | 2) + varint(len(value)) + value


def reference(field, target):
    return blob(field, integer(1, target))


def record(identity, kind, value):
    header = integer(1, identity) + blob(2, integer(1, kind) + integer(3, len(value)))
    return varint(len(header)) + header + value


parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("output", type=Path)
args = parser.parse_args()
cell = bytearray(20)
cell[0], cell[1], cell[8] = 5, 2, 2
cell[12:] = struct.pack("<d", 42)
row = integer(1, 0) + integer(2, 1) + blob(6, bytes(cell)) + blob(7, b"\0\0")
store = blob(3, blob(1, integer(1, 0) + reference(2, 12)))
records = b"".join([
    record(1, 1, reference(1, 2)),
    record(2, 2, blob(1, b"Sheet") + reference(2, 10)),
    record(10, 6000, reference(2, 11)),
    record(11, 6001, integer(6, 1) + integer(7, 1) + blob(4, store)),
    record(12, 6002, integer(1, 1) + integer(2, 0) + integer(3, 1) + integer(4, 1)
           + blob(5, row) + integer(6, 5) + integer(7, 1)),
])
length = len(records) - 1
literal = bytes([length << 2]) if len(records) <= 60 else bytes([60 << 2, length])
block = varint(len(records)) + literal + records
framed = b"\0" + len(block).to_bytes(3, "little") + block
args.output.parent.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(args.output, "w") as package:
    entry = zipfile.ZipInfo("Index/Document.iwa", date_time=(1980, 1, 1, 0, 0, 0))
    entry.create_system = 3  # Keep the archive byte-identical across producer platforms.
    entry.external_attr = 0o100644 << 16
    package.writestr(entry, framed)
