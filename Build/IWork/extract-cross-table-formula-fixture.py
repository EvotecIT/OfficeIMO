"""Extract the sixteen rectangular cross-table cases from the pinned native fixture.

Requires opt-in numbers-parser 4.19.0. The native package is never rewritten.
"""
import argparse
import hashlib
import json
import numbers
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.xrefs import xl_col_to_name, xl_rowcol_to_cell

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("source", type=Path)
parser.add_argument("output", type=Path)
parser.add_argument("--whole-axis-output", type=Path, help="Also reproduce the native whole-axis/body-range cases.")
args = parser.parse_args()
if version("numbers-parser") != "4.19.0":
    raise RuntimeError("The extractor requires numbers-parser 4.19.0.")
source_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
expected_hash = "9371c5b1d6ee4dfa17569097f064eba9c67f804d88b48638efbbeeb459d07dd4"
if source_hash != expected_hash:
    raise RuntimeError("The source differs from the pinned upstream fixture.")
document = Document(str(args.source))
sheet = document.sheets["Main Sheet"]
table = sheet.tables["Formula Tests"]
model = document._model
catalog = model.objects[model.objects[table._table_id].base_data_store.formula_table.identifier]
formulas = {entry.key: entry.formula for entry in catalog.entries}
cases = []
for row in range(25, 41):
    cell = table.cell(row, 0)
    nodes = formulas[cell._formula_id].AST_node_array.AST_node
    reference_nodes = [node for node in nodes if node.HasField("AST_cross_table_reference_extra_info")]
    assert len(reference_nodes) == 1 and reference_nodes[0].AST_node_type == 67
    target_uuid = NumbersUUID(reference_nodes[0].AST_cross_table_reference_extra_info.table_id).hex
    target_id = model.table_uuids_to_id(target_uuid)
    assert target_id is not None
    target_sheet = next(s for s in document.sheets if any(t._table_id == target_id for t in s.tables))
    target_table = next(t for t in target_sheet.tables if t._table_id == target_id)
    assert cell.value == 6.0
    assert cell.formula.startswith("COUNTA(Food Table::") and cell.formula.endswith(")")
    reference_address = cell.formula[len("COUNTA(Food Table::"):-1]
    cases.append({"sourceSheet": sheet.name, "sourceTable": table.name,
                  "row": row + 1, "column": 1, "sourceFormula": cell.formula,
                  "cachedValue": cell.value, "referenceAddress": reference_address, "targetSheet": target_sheet.name,
                  "targetTable": target_table.name, "targetUuid": target_uuid,
                  "nodeType": 67})
manifest = {"upstream": "https://github.com/masaccio/numbers-parser",
            "revision": "1c6c5c3d2e29a9abb601596678089f0a6c85d64c",
            "upstreamPath": "tests/data/create-formulas.numbers",
            "sourceSha256": source_hash, "extractorVersion": "numbers-parser 4.19.0",
            "license": "MIT, copyright Jon Connell",
            "qualification": "Independent native reference metadata and cached values; no Apple export or rendering oracle.",
            "cases": cases}
args.output.write_text(json.dumps(manifest, indent=2) + "\n")
print(f"Extracted {len(cases)} cases from {source_hash}")

if args.whole_axis_output:
    references = sheet.tables["Reference Tests"]
    ast = model.formula_ast(references._table_id)
    targets = {t._table_id: (s, t) for s in document.sheets for t in s.tables}
    axis_cases = []
    for row in range(1, references.num_rows):
        cell = references.cell(row, 0)
        nodes = ast[cell._formula_id]
        assert len(nodes) == 2 and nodes[1].AST_node_type == 16
        node = nodes[0]
        assert node.AST_node_type in (36, 67) and node.HasField("AST_cross_table_reference_extra_info")
        reference = model.node_to_ref(references._table_id, row, 0, node)
        assert (reference.row_start is None) != (reference.col_start is None)
        target_sheet, target = targets[reference.to_table_id]
        footer_rows = model.objects[target._table_id].number_of_footer_rows
        if reference.row_start is None:
            first_row, last_row = target.num_header_rows, target.num_rows - footer_rows - 1
            first_column = reference.col_start
            last_column = reference.col_end if reference.col_end is not None else first_column
            first_row_absolute = last_row_absolute = True
            first_column_absolute = reference.col_start_is_abs
            last_column_absolute = reference.col_end_is_abs if reference.col_end is not None else first_column_absolute
            first = xl_col_to_name(first_column, col_abs=first_column_absolute)
            last = xl_col_to_name(last_column, col_abs=last_column_absolute)
        else:
            first_column, last_column = target.num_header_cols, target.num_cols - 1
            first_row = reference.row_start
            last_row = reference.row_end if reference.row_end is not None else first_row
            first_column_absolute = last_column_absolute = True
            first_row_absolute = reference.row_start_is_abs
            last_row_absolute = reference.row_end_is_abs if reference.row_end is not None else first_row_absolute
            first = ("$" if first_row_absolute else "") + str(first_row + 1)
            last = ("$" if last_row_absolute else "") + str(last_row + 1)
        values = [target.cell(r, c).value for r in range(first_row, last_row + 1)
                  for c in range(first_column, last_column + 1)]
        function = {30: "COUNT", 31: "COUNTA", 168: "SUM"}[nodes[1].AST_function_node_index]
        numeric = [v for v in values if isinstance(v, numbers.Number) and not isinstance(v, bool)]
        computed = sum(v is not None for v in values) if function == "COUNTA" else len(numeric) if function == "COUNT" else sum(numeric)
        assert computed == cell.value, (row + 1, cell.formula, computed, cell.value)
        first_cell = xl_rowcol_to_cell(first_row, first_column, row_abs=first_row_absolute, col_abs=first_column_absolute)
        last_cell = xl_rowcol_to_cell(last_row, last_column, row_abs=last_row_absolute, col_abs=last_column_absolute)
        axis_cases.append({"sourceSheet": sheet.name, "sourceTable": references.name,
                           "row": row + 1, "column": 1, "sourceFormula": cell.formula,
                           "function": function, "cachedValue": cell.value,
                           "sourceCoordinateRange": first + ":" + last,
                           "destinationBodyRange": first_cell if first_cell == last_cell else first_cell + ":" + last_cell,
                           "targetSheet": target_sheet.name, "targetTable": target.name,
                           "targetRows": target.num_rows, "targetColumns": target.num_cols,
                           "headerRows": target.num_header_rows, "headerColumns": target.num_header_cols,
                           "footerRows": footer_rows, "computedCurrentBodyValue": computed,
                           "nodeType": node.AST_node_type})
    assert len(axis_cases) == 59
    axis_manifest = {key: value for key, value in manifest.items() if key not in ("cases", "qualification")}
    axis_manifest["qualification"] = "Independent whole-axis coordinates and table-body values agree with native caches; fixed XLSX body ranges do not preserve named labels or automatic table expansion. No Apple export/rendering oracle."
    axis_manifest["cases"] = axis_cases
    args.whole_axis_output.write_text(json.dumps(axis_manifest, indent=2) + "\n")
    print(f"Extracted {len(axis_cases)} whole-axis cases with matching current-body values")
