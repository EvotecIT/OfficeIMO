"""Create a currency-format corpus with numbers-parser 4.19.0 (opt-in only)."""
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

# Source display comes from the independent producer. Destination display is the
# explicitly chosen portable identifier-prefix contract, not a native Apple oracle.
cases = [
    ("USD grouped", 1234.5, "USD", 2, True, 0, False, "USD 1,234.50"),
    ("USD minus", -1234.5, "USD", 2, True, 0, False, "USD -1,234.50"),
    ("EUR red", -12.5, "EUR", 2, False, 1, False, "EUR 12.50"),
    ("GBP parentheses", -12.5, "GBP", 1, False, 2, False, "GBP (12.5)"),
    ("JPY red parentheses", -1234.625, "JPY", 0, True, 3, False, "JPY (1,235)"),
    ("PLN precision", 12.125, "PLN", 3, False, 0, False, "PLN 12.125"),
    ("CHF zero", 0.0, "CHF", 2, False, 0, False, "CHF 0.00"),
    ("USD accounting positive", 1234.5, "USD", 2, True, 0, True, "USD 1,234.50"),
    ("USD accounting negative", -1234.5, "USD", 2, True, 0, True, "USD (1,234.50)"),
    ("EUR accounting zero", 0.0, "EUR", 2, False, 0, True, "EUR 0.00"),
    ("USD automatic", 1.5, "USD", 253, False, 0, False, "USD 1.5"),
]
document = Document(sheet_name="Currency formats", table_name="Formats", num_rows=len(cases),
                    num_cols=2, num_header_rows=0, num_header_cols=0)
table = document.sheets[0].tables[0]
table.col_width(0, 200.0)
table.col_width(1, 125.0)
for row, (label, value, code, decimals, grouping, negative, accounting, _) in enumerate(cases):
    table.row_height(row, 22.0)
    table.write(row, 0, label)
    table.write(row, 1, value)
    table.set_cell_formatting(row, 1, "currency", currency_code=code, decimal_places=decimals,
                              show_thousands_separator=grouping,
                              negative_style=NegativeNumberStyle(negative),
                              use_accounting_style=accounting)
document.save(str(args.output))

reopened = Document(str(args.output)).sheets[0].tables[0]
expected = []
for row, (label, value, code, decimals, grouping, negative, accounting, destination) in enumerate(cases):
    cell = reopened.cell(row, 1)
    assert reopened.cell(row, 0).value == label
    assert math.isclose(cell.value, value, rel_tol=1e-12, abs_tol=1e-12)
    # Version-specific protobuf inspection remains in opt-in fixture tooling.
    native = cell._model.table_format(cell._table_id, cell._currency_format_id)
    assert native.format_type == 257 and native.currency_code == code
    assert native.decimal_places == decimals and native.negative_style == negative
    assert native.show_thousands_separator == grouping and native.use_accounting_style == accounting
    assert cell._flags & (1 << 14)
    expected.append({
        "row": row + 1, "label": label, "value": cell.value,
        "currencyCode": code, "decimalPlaces": None if decimals == 253 else decimals,
        "thousandsSeparator": grouping, "negativeStyle": negative,
        "useAccountingStyle": accounting, "sourceDisplayText": cell.formatted_value,
        "destinationDisplayText": destination,
    })
manifest = {
    "producer": "numbers-parser", "producerVersion": "4.19.0",
    "sourceSha256": hashlib.sha256(args.output.read_bytes()).hexdigest(),
    "qualification": "Independent-producer metadata and source display text. Destination display uses explicit currency identifiers with reported approximation; no Apple native export or appearance claim.",
    "cases": expected,
}
args.output.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print(json.dumps(manifest, indent=2))
