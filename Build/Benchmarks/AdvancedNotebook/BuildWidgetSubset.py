"""Retain the linked Play/slider view and its dependencies in Widget List."""

import hashlib
import json
from pathlib import Path
import sys

STATE_MIME = "application/vnd.jupyter.widget-state+json"
VIEW_MIME = "application/vnd.jupyter.widget-view+json"
SOURCE_HASH = "f19ad2075cc57b504ef2d83cc9f0f19cf025d108e880d1e5d8c9c6940501c9dd"
PLAY_ROW = "3f71a71a7af341a2a69f423d4cbfa008"
PLAY_LINK = "eddc968a6e184d599b89553aa7bc7ef9"


def references(value):
    if isinstance(value, str) and value.startswith("IPY_MODEL_"):
        yield value[len("IPY_MODEL_"):]
    elif isinstance(value, dict):
        for child in value.values():
            yield from references(child)
    elif isinstance(value, list):
        for child in value:
            yield from references(child)


def main(source, destination):
    raw = Path(source).read_bytes()
    if hashlib.sha256(raw).hexdigest() != SOURCE_HASH:
        raise ValueError("Source must be the pinned Jupyter nbconvert Widget List fixture.")
    notebook = json.loads(raw)
    notebook["cells"] = [
        cell for cell in notebook["cells"]
        if any(output.get("data", {}).get(VIEW_MIME, {}).get("model_id") == PLAY_ROW
               for output in cell.get("outputs", []))
    ]
    if len(notebook["cells"]) != 1:
        raise ValueError("Expected one saved Play row.")

    envelope = notebook["metadata"]["widgets"][STATE_MIME]
    original = envelope["state"]
    selected = {}

    def visit(model_id):
        if model_id in selected:
            return
        selected[model_id] = original[model_id]
        for dependency in references(selected[model_id]):
            visit(dependency)

    visit(PLAY_ROW)
    visit(PLAY_LINK)
    envelope["state"] = selected
    document = json.dumps(notebook, indent=1, ensure_ascii=False, sort_keys=True)
    Path(destination).write_bytes((document + "\n").encode("utf-8"))


if __name__ == "__main__":
    main(*sys.argv[1:])
