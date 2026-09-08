import csv
import hashlib
import json
import sys
from pathlib import Path

path = Path(sys.argv[1])
rows = fields = characters = nonempty = estimate = cached_fields = cached_bytes = 0
with path.open(encoding="utf-8-sig", newline="") as stream:
    reader = csv.reader(stream)
    headers = next(reader)
    for row in reader:
        assert len(row) == 14
        rows += 1
        for value in row:
            length = len(value.encode("utf-16-le")) // 2
            fields += 1
            characters += length
            if length:
                nonempty += 1
                estimate += (22 + 2 * length + 7) & ~7
                if length == 1 and ord(value) < 128:
                    cached_fields += 1
                    cached_bytes += 24
assert (rows, fields, characters) == (65535, 917490, 7253195)
print(json.dumps({
    "fixture": path.name,
    "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    "rows": rows, "fields": fields, "utf16Characters": characters,
    "nonemptyFields": nonempty, "utf16PayloadBytes": characters * 2,
    "freshStringAllocationEstimateBytes": estimate,
    "reusedSingleAsciiCharacterFields": cached_fields,
    "reusedStringAllocationBytes": cached_bytes,
    "actualFreshFieldStringAllocationEstimateBytes": estimate - cached_bytes,
    "assumption": "One fresh x64 .NET string per nonempty field, 22-byte base including terminator, 8-byte alignment. Not a lower bound for implementations that reuse immutable values. Excludes headers and parser state."
}, indent=2))
