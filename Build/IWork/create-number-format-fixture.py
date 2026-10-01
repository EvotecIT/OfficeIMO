"""Create a number-format corpus with numbers-parser 4.19.0 (opt-in only)."""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from pathlib import Path

from numbers_parser import Document, NegativeNumberStyle

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("output", type=Path)
args = parser.parse_args()
if version("numbers-parser") != "4.19.0":
    raise RuntimeError("The fixture requires numbers-parser 4.19.0.")

# Values avoid rounding ties so the corpus tests formatting, not rounding-policy
# disagreements between producers. Both percentage scaling and negative sections
# must operate on the stored numeric value, without replacing it with text.
cases = [
    ("Grouped number", 1234.567, "number", 2, True, NegativeNumberStyle.MINUS),
    ("Negative minus", -1234.567, "number", 2, True, NegativeNumberStyle.MINUS),
    ("Negative red", -12.3, "number", 3, False, NegativeNumberStyle.RED),
    ("Negative parentheses", -12.34, "number", 1, False, NegativeNumberStyle.PARENTHESES),
    ("Negative red parentheses", -12.34, "number", 0, False, NegativeNumberStyle.RED_AND_PARENTHESES),
    ("Percentage decimals", 0.125, "percentage", 2, False, NegativeNumberStyle.MINUS),
    ("Grouped percentage", 12.3456, "percentage", 1, True, NegativeNumberStyle.MINUS),
    ("Percentage red", -0.5, "percentage", 0, False, NegativeNumberStyle.RED),
    ("Percentage parentheses", -0.125, "percentage", 2, False, NegativeNumberStyle.PARENTHESES),
    ("Percentage red parentheses", -0.125, "percentage", 2, False, NegativeNumberStyle.RED_AND_PARENTHESES),
    ("Zero percentage", 0.0, "percentage", 0, False, NegativeNumberStyle.MINUS),
    ("Automatic number", 1.5, "number", None, False, NegativeNumberStyle.MINUS),
    ("Automatic percentage", 0.125, "percentage", None, False, NegativeNumberStyle.MINUS),
]
document = Document(sheet_name="Number formats", table_name="Formats", num_rows=len(cases),
                    num_cols=2, num_header_rows=0, num_header_cols=0)
table = document.sheets[0].tables[0]
table.col_width(0, 190.0)
table.col_width(1, 110.0)
for row, (label, value, kind, decimals, grouping, negative) in enumerate(cases):
    table.row_height(row, 22.0)
    table.write(row, 0, label)
    table.write(row, 1, value)
    table.set_cell_formatting(row, 1, kind, decimal_places=decimals,
                              show_thousands_separator=grouping, negative_style=negative)
document.save(str(args.output))

reopened = Document(str(args.output)).sheets[0].tables[0]
expected = []
for row, (label, value, kind, decimals, grouping, negative) in enumerate(cases):
    cell = reopened.cell(row, 1)
    assert reopened.cell(row, 0).value == label
    assert math.isclose(cell.value, value, rel_tol=1e-12, abs_tol=1e-12), (label, cell.value, value)
    # The pinned producer exposes display text publicly. Its protobuf metadata
    # inspection is private and version-specific; it stays in opt-in tooling.
    native = cell._model.table_format(cell._table_id, cell._num_format_id)
    assert native.format_type == (258 if kind == "percentage" else 256)
    assert native.decimal_places == (253 if decimals is None else decimals)
    assert native.show_thousands_separator == grouping
    assert native.negative_style == negative.value
    assert cell._buffer[0] == 5 and cell._flags & 1
    raw = cell._buffer[12:28]
    coefficient = int.from_bytes(raw[:14], "little") + ((raw[14] & 1) << 112)
    exponent = (((raw[15] & 0x7f) << 7) | (raw[14] >> 1)) - 0x1820
    if coefficient:
        while coefficient % 10 == 0:
            coefficient //= 10
            exponent += 1
        source_text = ("-" if raw[15] & 0x80 else "") + str(coefficient) + "E" + str(exponent)
    else:
        source_text = "0"
    expected.append({
        "row": row + 1, "label": label, "value": cell.value,
        "kind": "Percentage" if kind == "percentage" else "Number",
        "decimalPlaces": decimals, "thousandsSeparator": native.show_thousands_separator,
        "negativeStyle": native.negative_style, "displayText": cell.formatted_value,
        "sourceNumberText": source_text, "numericValueIsApproximate": len(str(coefficient)) > 15,
    })
manifest = {
    "producer": "numbers-parser", "producerVersion": "4.19.0",
    "sourceSha256": hashlib.sha256(args.output.read_bytes()).hexdigest(),
    "qualification": "Independent-producer semantics and display text; no Apple native export or appearance claim.",
    "cases": expected,
}
args.output.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print(json.dumps(manifest, indent=2))
