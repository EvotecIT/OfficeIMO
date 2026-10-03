"""Create the independent sizing fixture with numbers-parser 4.19.0 (opt-in only)."""
import argparse
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("output", type=Path)
args = parser.parse_args()
if version("numbers-parser") != "4.19.0":
    raise RuntimeError("The fixture requires numbers-parser 4.19.0.")

document = Document(sheet_name="Dimensions", table_name="Nonuniform", num_rows=3,
                    num_cols=2, num_header_rows=0, num_header_cols=0)
table = document.sheets[0].tables[0]
for row, height in enumerate((20.0, 10.0, 30.0)):
    table.row_height(row, height)
    table.write(row, 0, f"Row {row + 1}")
for column, width in enumerate((40.0, 20.0)):
    table.col_width(column, width)
document.save(str(args.output))

# Reopen through the independent producer as well as OfficeIMO's corpus tests.
reopened = Document(str(args.output)).sheets[0].tables[0]
assert [reopened.row_height(row) for row in range(3)] == [20.0, 10.0, 30.0]
assert [reopened.col_width(column) for column in range(2)] == [40.0, 20.0]
