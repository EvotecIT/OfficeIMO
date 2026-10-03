"""Create a scientific-format corpus with numbers-parser 4.19.0 (opt-in only)."""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from pathlib import Path

from numbers_parser import Document
from fixture_number_values import source_number

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("output", type=Path)
args = parser.parse_args()
if version("numbers-parser") != "4.19.0":
    raise RuntimeError("The fixture requires numbers-parser 4.19.0.")

# Explicit cases avoid rounding ties and use the producer's reopened display text.
# The automatic sentinel is a metadata oracle only: this producer prints 253
# fractional digits for it. Its display is not an Apple automatic-format oracle.
cases = [
    ("Large value", 12345.6, 3, None),
    ("Small value", 0.0000123456, 4, None),
    ("Negative value", -987654.0, 2, None),
    ("Zero", 0.0, 2, None),
    ("No decimals", 1.0, 0, None),
    ("Negative unit", -1.0, 1, None),
    ("Large exponent", 6.25e60, 6, None),
    ("Three-digit exponent", 1.25e-100, 3, None),
    ("Thirty decimals", 0.125, 30, None),
    ("Zero thirty decimals", 0.0, 30, None),
    ("Exponent rollover", 999.9, 2, None),
    ("Negative fraction", -0.0625, 3, None),
    ("Automatic unit", 1.5, None, "1.5E+00"),
    ("Automatic small", 0.00125, None, "1.25E-03"),
]
document = Document(sheet_name="Scientific formats", table_name="Formats", num_rows=len(cases),
                    num_cols=2, num_header_rows=0, num_header_cols=0)
table = document.sheets[0].tables[0]
table.col_width(0, 185.0)
table.col_width(1, 235.0)
for row, (label, value, decimals, _) in enumerate(cases):
    table.row_height(row, 22.0)
    table.write(row, 0, label)
    table.write(row, 1, value)
    table.set_cell_formatting(row, 1, "scientific", decimal_places=decimals)
document.save(str(args.output))

reopened = Document(str(args.output)).sheets[0].tables[0]
expected = []
for row, (label, value, decimals, portable_automatic) in enumerate(cases):
    cell = reopened.cell(row, 1)
    assert reopened.cell(row, 0).value == label
    assert math.isclose(cell.value, value, rel_tol=1e-12, abs_tol=1e-112)
    native = cell._model.table_format(cell._table_id, cell._num_format_id)
    assert native.format_type == 259
    assert native.decimal_places == (253 if decimals is None else decimals)
    assert {field.number for field, _ in native.ListFields()} == {1, 2}
    assert cell._flags & (1 << 13) and not cell._flags & (1 << 14)
    source_text, portable_value, approximate = source_number(cell)
    expected.append({
        "row": row + 1, "label": label, "value": cell.value,
        "decimalPlaces": decimals, "sourceDisplayText": cell.formatted_value,
        "sourceDisplayIsQualified": decimals is not None,
        "destinationDisplayText": cell.formatted_value if decimals is not None else portable_automatic,
        "sourceNumberText": source_text, "portableValue": portable_value,
        "numericValueIsApproximate": approximate,
    })
manifest = {
    "producer": "numbers-parser", "producerVersion": "4.19.0",
    "sourceSha256": hashlib.sha256(args.output.read_bytes()).hexdigest(),
    "qualification": "Independent-producer explicit scientific precision and automatic-mode metadata. No Apple native export or appearance claim.",
    "automaticDisplayGap": "The pinned producer interprets sentinel253 as253 fractional digits. Its automatic display text is unqualified. Destination automatic text uses the separately reported portable optional-decimal approximation.",
    "numericValueContract": "value records the producer's floating-point decoder result. portableValue uses Python Decimal on exact source storage text before float conversion; it avoids the producer's intermediate floating-point arithmetic. OfficeIMO preserves that portable value.",
    "cases": expected,
}
args.output.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print(json.dumps(manifest, indent=2))
