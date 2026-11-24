"""Preserve the exact producer's MIME cells and add selectable text witnesses."""
import argparse
import copy
import hashlib
from pathlib import Path

import nbformat

parser = argparse.ArgumentParser()
parser.add_argument("source", type=Path)
parser.add_argument("output", type=Path)
parser.add_argument("--case", action="append", nargs=2, metavar=("CELL_ID", "WITNESS"), required=True)
args = parser.parse_args()
args.output.mkdir(exist_ok=True, parents=True)
source = nbformat.read(args.source, as_version=4)
by_id = {cell.id: cell for cell in source.cells}
for name, witness in args.case:
    cell = copy.deepcopy(by_id[name])
    assert cell.outputs[0].data.keys() == {"application/vnd.plotly.v1+json"}
    note = nbformat.v4.new_markdown_cell(source="# " + witness, id="heading-" + name)
    notebook = nbformat.v4.new_notebook(cells=[note, cell], metadata=copy.deepcopy(source.metadata))
    nbformat.validate(notebook)
    output = args.output / (name + ".ipynb")
    nbformat.write(notebook, output)
    print(hashlib.sha256(output.read_bytes()).hexdigest(), output)
