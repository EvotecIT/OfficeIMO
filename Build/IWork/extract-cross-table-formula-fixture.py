"""Extract the sixteen rectangular cross-table cases from the pinned native fixture.

Requires opt-in numbers-parser 4.19.0. The native package is never rewritten.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document
from numbers_parser.numbers_uuid import NumbersUUID

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("source", type=Path)
parser.add_argument("output", type=Path)
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
    cases.append({"sourceSheet": sheet.name, "sourceTable": table.name,
                  "row": row + 1, "column": 1, "sourceFormula": cell.formula,
                  "cachedValue": cell.value, "targetSheet": target_sheet.name,
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
