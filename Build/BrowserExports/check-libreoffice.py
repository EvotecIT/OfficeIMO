"""Check LibreOffice's independently imported values; record its early-date limitation."""
import csv
import json
import pathlib
import subprocess
import sys

if len(sys.argv) != 2:
    raise SystemExit("Usage: python check-libreoffice.py <evidence-directory-containing-rich-auto.csv>")
root = pathlib.Path(sys.argv[1])
with (root / "rich-auto.csv").open(encoding="utf-8", newline="") as stream:
    rows = list(csv.reader(stream))
assert rows[1][0].replace("\r\n", "\n") == "DC<&\"'\n\t🧪שלום", rows[1][0]
assert rows[1][1] == "Łódź" and rows[1][3:] == ["12.50", "TRUE"]
assert rows[2][0] == "=literal" and rows[3][0] == "_x0041_"
assert rows[1][2] == "2026-10-05 12:34" and rows[3][2] == "1900-03-01 00:00"
report = {
    "application": subprocess.check_output(["libreoffice", "--version"], text=True).strip(),
    "valuesPassed": True,
    "modernAndMarch1900DatesPassed": True,
    "crLfPreserved": "\r\n" in rows[1][0],
    "early1900DatesPassed": rows[2][2] == "1900-02-28 12:00" and rows[4][2] == "1900-01-01 00:00",
    "early1900Actual": [rows[2][2], rows[4][2]],
    "early1900Expected": ["1900-02-28 12:00", "1900-01-01 00:00"],
    "note": "Excel and OfficeIMO verify serials 59.5 and 1. LibreOffice 24.2 displays dates before March 1900 one day early."
}
(root / "spot-check.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
print(json.dumps(report, indent=2))
