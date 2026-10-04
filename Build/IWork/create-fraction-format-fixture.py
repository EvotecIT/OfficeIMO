"""Create a fraction-format corpus with numbers-parser 4.19.0 (opt-in only)."""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from pathlib import Path

from numbers_parser import Document, FractionAccuracy
from fixture_number_values import source_number

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("output", type=Path)
args = parser.parse_args()
if version("numbers-parser") != "4.19.0":
    raise RuntimeError("The fixture requires numbers-parser 4.19.0.")

# None means that the reopened public display is the independent oracle.
# Explicit alternatives record a destination contract where this producer loses
# a negative whole part, does not normalize rollover, or rounds midpoint ties
# differently. These alternatives are not independent Apple display evidence.
cases = [
    ("One digit", 0.2578125, FractionAccuracy.ONE, None),
    ("Two digits", 0.2578125, FractionAccuracy.TWO, None),
    ("Three digits", 0.2578125, FractionAccuracy.THREE, None),
    ("Halves", 0.625, FractionAccuracy.HALVES, None),
    ("Quarters", 0.6875, FractionAccuracy.QUARTERS, None),
    ("Eighths", 0.65625, FractionAccuracy.EIGHTHS, None),
    ("Sixteenths", 0.375, FractionAccuracy.SIXTEENTHS, None),
    ("Tenths", 0.3125, FractionAccuracy.TENTHS, None),
    ("Hundredths", 0.3125, FractionAccuracy.HUNDREDTHS, None),
    ("Negative proper", -0.375, FractionAccuracy.EIGHTHS, None),
    ("Mixed", 1.375, FractionAccuracy.EIGHTHS, None),
    ("Negative mixed", -1.375, FractionAccuracy.EIGHTHS, "-1 3/8"),
    ("Zero", 0.0, FractionAccuracy.THREE, None),
    ("Whole", 2.0, FractionAccuracy.TWO, None),
    ("Negative whole", -2.0, FractionAccuracy.HALVES, "-2"),
    ("Fixed rollover", 1.99, FractionAccuracy.HALVES, "2"),
    ("Variable rollover", 1.99, FractionAccuracy.ONE, "2"),
    ("Positive half tie", 0.25, FractionAccuracy.HALVES, "1/2"),
    ("Negative half tie", -0.25, FractionAccuracy.HALVES, "-1/2"),
    # The writer stores this value just below .625; it is not an exact tie.
    ("Near quarter tie", 0.625, FractionAccuracy.QUARTERS, None),
    ("Hundredth tie", 0.125, FractionAccuracy.HUNDREDTHS, "13/100"),
    ("Tiny fraction", 0.00001, FractionAccuracy.THREE, None),
    ("Large mixed", 12345.375, FractionAccuracy.EIGHTHS, None),
    ("Negative three digits", -0.2578125, FractionAccuracy.THREE, None),
    ("Beyond Int64", 1e20, FractionAccuracy.EIGHTHS, None),
    ("Negative beyond Int64", -1e20, FractionAccuracy.EIGHTHS, "-100000000000000000000"),
]
names = {
    FractionAccuracy.ONE: "OneDigitDenominator", FractionAccuracy.TWO: "TwoDigitDenominator",
    FractionAccuracy.THREE: "ThreeDigitDenominator", FractionAccuracy.HALVES: "Halves",
    FractionAccuracy.QUARTERS: "Quarters", FractionAccuracy.EIGHTHS: "Eighths",
    FractionAccuracy.SIXTEENTHS: "Sixteenths", FractionAccuracy.TENTHS: "Tenths",
    FractionAccuracy.HUNDREDTHS: "Hundredths",
}
codes = {
    FractionAccuracy.ONE: "# ?/?", FractionAccuracy.TWO: "# ??/??",
    FractionAccuracy.THREE: "# ???/???", FractionAccuracy.HALVES: "# ?/2",
    FractionAccuracy.QUARTERS: "# ?/4", FractionAccuracy.EIGHTHS: "# ?/8",
    FractionAccuracy.SIXTEENTHS: "# ??/16", FractionAccuracy.TENTHS: "# ??/10",
    FractionAccuracy.HUNDREDTHS: "# ???/100",
}
codes = {accuracy: code + ";-" + code + ";0" for accuracy, code in codes.items()}
document = Document(sheet_name="Fraction formats", table_name="Formats", num_rows=len(cases),
                    num_cols=2, num_header_rows=0, num_header_cols=0)
table = document.sheets[0].tables[0]
table.col_width(0, 185.0)
table.col_width(1, 150.0)
for row, (label, value, accuracy, _) in enumerate(cases):
    table.row_height(row, 22.0)
    table.write(row, 0, label)
    table.write(row, 1, value)
    table.set_cell_formatting(row, 1, "fraction", fraction_accuracy=accuracy)
document.save(str(args.output))

reopened = Document(str(args.output)).sheets[0].tables[0]
expected = []
for row, (label, value, accuracy, destination) in enumerate(cases):
    cell = reopened.cell(row, 1)
    assert reopened.cell(row, 0).value == label
    assert math.isclose(cell.value, value, rel_tol=1e-12, abs_tol=1e-15)
    native = cell._model.table_format(cell._table_id, cell._num_format_id)
    assert native.format_type == 262 and native.fraction_accuracy == int(accuracy)
    assert {field.number for field, _ in native.ListFields()} == {1, 11}
    assert cell._flags & (1 << 13) and not cell._flags & (1 << 14)
    source_text, portable_value, approximate = source_number(cell)
    expected.append({
        "row": row + 1, "label": label, "value": cell.value, "portableValue": portable_value,
        "sourceNumberText": source_text, "numericValueIsApproximate": approximate,
        "sourceAccuracy": int(accuracy), "fractionAccuracy": names[accuracy],
        "destinationFormatCode": codes[accuracy], "sourceDisplayText": cell.formatted_value,
        "sourceDisplayIsQualified": destination is None,
        "destinationDisplayText": cell.formatted_value if destination is None else destination,
    })
manifest = {
    "producer": "numbers-parser", "producerVersion": "4.19.0",
    "sourceSha256": hashlib.sha256(args.output.read_bytes()).hexdigest(),
    "qualification": "Independent-producer nine-mode fraction metadata and non-midpoint display examples. No Apple native export or appearance claim.",
    "displayGaps": "The producer drops negative whole parts and does not normalize some rollover. Fixed-denominator midpoint ties use Python ties-to-even. Explicit destination alternatives qualify OfficeIMO's reported approximation, not producer or Apple display equivalence.",
    "numericValueContract": "value retains the producer decoder result; portableValue derives a correctly rounded float from exact stored Decimal128 text through Python Decimal.",
    "cases": expected,
}
args.output.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print(json.dumps(manifest, indent=2))
