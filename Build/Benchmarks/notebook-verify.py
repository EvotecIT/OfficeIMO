"""Independent semantic and artifact checks used after PowerForge operations."""

import argparse
import base64
import hashlib
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import subprocess
import xml.etree.ElementTree as ET


def multiline(value):
    return "".join(value) if isinstance(value, list) else value


def notebook_sources(notebook):
    if "cells" in notebook:
        return [[cell["cell_type"], multiline(cell.get("source", ""))] for cell in notebook["cells"]]
    result = []
    for worksheet in notebook["worksheets"]:
        for cell in worksheet["cells"]:
            kind = cell["cell_type"]
            if kind == "heading":
                title = multiline(cell.get("source", ""))
                result.append(["markdown", "#" * cell.get("level", 1) + " " + title])
            else:
                key = "input" if kind == "code" else "source"
                result.append([kind, multiline(cell.get(key, ""))])
    return result


def bundle(data):
    return {mime: multiline(value) if mime.startswith(("text/", "image/")) or mime == "application/javascript" else value
            for mime, value in data.items()}


def normalize_v4(notebook):
    for cell in notebook["cells"]:
        if "source" in cell:
            cell["source"] = multiline(cell["source"])
        for name, attachment in cell.get("attachments", {}).items():
            cell["attachments"][name] = bundle(attachment)
        for output in cell.get("outputs", []):
            if output["output_type"] == "stream":
                output["text"] = multiline(output["text"])
            if "data" in output:
                output["data"] = bundle(output["data"])
    return notebook


def deviations(source, result, path="$"):
    if isinstance(source, dict) and isinstance(result, dict):
        for field in sorted(source.keys() - result.keys()):
            yield {"path": f"{path}.{field}", "change": "missing"}
        for field in sorted(result.keys() - source.keys()):
            yield {"path": f"{path}.{field}", "change": "added"}
        for field in sorted(source.keys() & result.keys()):
            yield from deviations(source[field], result[field], f"{path}.{field}")
    elif isinstance(source, list) and isinstance(result, list):
        if len(source) != len(result):
            yield {"path": path, "change": "length", "expected": len(source), "actual": len(result)}
        for index, (expected, actual) in enumerate(zip(source, result)):
            yield from deviations(expected, actual, f"{path}[{index}]")
    elif type(source) is not type(result) or source != result:
        yield {"path": path, "change": "value"}


class HtmlWitness(HTMLParser):
    def __init__(self):
        super().__init__()
        self.images = []
        self.tables = 0
        self.text = []
        self.skipped = ""

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skipped = tag
        if tag == "img":
            self.images.append(dict(attrs).get("src", ""))
        self.tables += tag == "table"

    def handle_endtag(self, tag):
        if tag == self.skipped:
            self.skipped = ""

    def handle_data(self, data):
        if not self.skipped:
            self.text.append(data)


def markers(notebook):
    for kind, source in notebook_sources(notebook):
        if kind == "markdown":
            yield re.search(r"NotebookCell\d+", source).group()
        elif kind == "code":
            yield re.search(r"CodeBegin\d+", source).group()
            yield re.search(r"CodeEnd\d+", source).group()
    for cell in notebook["cells"]:
        for output in cell.get("outputs", []):
            if output["output_type"] == "stream":
                stream_text = multiline(output["text"])
                yield re.search(r"StreamBegin\d+", stream_text).group()
                yield re.search(r"StreamEnd\d+", stream_text).group()


def check_pdf(output, text_tool, image_tool):
    result = subprocess.run([text_tool, "-bbox-layout", "-enc", "UTF-8", str(output), "-"],
                            check=True, capture_output=True)
    xml = ET.fromstring(result.stdout)
    pages = xml.findall(".//{*}page")
    if not pages:
        raise ValueError("Poppler could not find any PDF pages.")
    words = []
    for page in pages:
        width = float(page.attrib["width"])
        height = float(page.attrib["height"])
        for word in page.findall(".//{*}word"):
            if (float(word.attrib["xMin"]) < -0.5 or float(word.attrib["xMax"]) > width + 0.5
                    or float(word.attrib["yMin"]) < -0.5 or float(word.attrib["yMax"]) > height + 0.5):
                raise ValueError("A PDF word lies outside the page box.")
            words.append("".join(word.itertext()))
    listing = subprocess.run([image_tool, "-list", str(output)], check=True, capture_output=True, text=True)
    image_count = sum(bool(re.match(r"^\s*\d+\s+\d+\s+image\b", line)) for line in listing.stdout.splitlines())
    return "".join(words), len(pages), image_count


def verify_import(source_bytes, output_path, manifest_path, case_name, companion_path):
    import nbformat

    fixtures = json.loads(manifest_path.read_text(encoding="utf-8"))["cases"]
    expected = next(fixture for fixture in fixtures if fixture["name"] == case_name)
    output_bytes = output_path.read_bytes()
    notebook = json.loads(output_bytes)
    if [notebook.get("nbformat"), notebook.get("nbformat_minor")] != [4, 5]:
        raise ValueError("Import did not output nbformat 4.5.")
    cells = notebook["cells"]
    ids = [cell["id"] for cell in cells]
    if len(ids) != len(set(ids)):
        raise ValueError("Import produced duplicate cell IDs.")
    nbformat.validate(nbformat.from_dict(notebook))
    normalize_v4(notebook)
    companion_size = 0
    if expected.get("companion"):
        if companion_path is None:
            raise ValueError("The paired import requires its pinned companion.")
        companion_bytes = companion_path.read_bytes()
        companion_size = len(companion_bytes)
        if hashlib.sha256(companion_bytes).hexdigest() != expected["companionSha256"]:
            raise ValueError("The paired import used a different companion.")
        if notebook_sources(notebook) != notebook_sources(json.loads(companion_bytes)):
            raise ValueError("The paired import changed the notebook's ordered source.")
    if [cell["cell_type"] for cell in cells] != expected["cellKinds"]:
        raise ValueError(f"{case_name} changed expected cell order, count, or kinds.")
    outputs = sum(len(cell.get("outputs", [])) for cell in cells)
    if outputs != expected["outputCount"]:
        raise ValueError(f"{case_name} produced {outputs} saved outputs, expected {expected['outputCount']}.")

    for witness in expected["witnesses"]:
        value = notebook
        for part in witness["path"]:
            value = value[part]
        if "equals" in witness:
            matched = value == witness["equals"]
        elif "contains" in witness:
            matched = witness["contains"] in value
        else:
            matched = value.startswith(witness["startsWith"])
        if not matched:
            raise ValueError(f"{case_name} did not preserve the sample at {witness['path']}.")

    report_path = output_path.with_name("conversion-report.json")
    diagnostics = json.loads(report_path.read_bytes())
    missing = set(expected.get("reportCodes", [])) - {item["code"] for item in diagnostics}
    if missing:
        raise ValueError(f"{case_name} did not report expected findings: {sorted(missing)}")
    return {
        "inputSha256": hashlib.sha256(source_bytes).hexdigest(),
        "inputBytes": len(source_bytes),
        "companionInputBytes": companion_size,
        "outputBytes": len(output_bytes),
        "verifiedCells": len(cells),
        "verifiedPages": 0,
        "verifiedImages": sum("image/png" in output.get("data", {}) for cell in cells for output in cell.get("outputs", [])),
        "verifiedOutputs": outputs,
        "semanticDifferences": 0,
    }


def verify(operation, input_path, output_path, text_tool, image_tool):
    source_bytes = input_path.read_bytes()
    source = json.loads(source_bytes)
    source_cells = notebook_sources(source)
    output_bytes = output_path.read_bytes()
    result = {
        "inputSha256": hashlib.sha256(source_bytes).hexdigest(),
        "inputBytes": len(source_bytes),
        "outputBytes": len(output_bytes),
        "verifiedCells": len(source_cells),
        "verifiedPages": 0,
        "verifiedImages": 0,
        "semanticDifferences": 0,
    }
    if operation == "SourceProjection":
        if json.loads(output_bytes) != source_cells:
            raise ValueError("The common projection changed ordered cell kinds or cell source.")
    elif operation in ("Parse", "RoundTrip", "RichRoundTrip"):
        saved = json.loads(output_bytes)
        if notebook_sources(saved) != source_cells and operation != "RichRoundTrip":
            raise ValueError("The notebook changed ordered cell kinds or cell source.")
        if source["nbformat"] != 4:
            differences = [{"path": "$.nbformat", "change": "legacy normalization",
                            "expected": source["nbformat"], "actual": saved["nbformat"]}]
            if notebook_sources(saved) != source_cells:
                differences.append({"path": "$.cells", "change": "legacy conversion changed ordered kinds or source"})
        else:
            differences = list(deviations(normalize_v4(source), normalize_v4(saved)))
        result["semanticDifferences"] = len(differences)
        (output_path.parent / "fidelity.json").write_text(json.dumps(differences, indent=2), encoding="utf-8")
        if differences and operation != "RichRoundTrip":
            raise ValueError("The notebook lost or changed content. Inspect fidelity.json.")
    else:
        expected_images = [base64.b64decode(multiline(data["image/png"])) for cell in source["cells"]
                           for data in [*cell.get("attachments", {}).values(),
                                        *(output.get("data", {}) for output in cell.get("outputs", []))]
                           if "image/png" in data]
        if operation == "Pdf":
            text, pages, images = check_pdf(output_path, text_tool, image_tool)
            result.update(verifiedPages=pages, verifiedImages=images)
        elif operation == "Html":
            witness = HtmlWitness()
            witness.feed(output_bytes.decode("utf-8"))
            text = "".join(witness.text)
            result["verifiedImages"] = verify_images(witness.images, expected_images, output_path.parent)
            if expected_images and not witness.tables:
                raise ValueError("The saved HTML has no table.")
        elif operation == "Markdown":
            text = output_bytes.decode("utf-8")
            sources = re.findall(r"!\[[^\]]*\]\(<?([^\s)>]+)>?[^)]*\)", text)
            result["verifiedImages"] = verify_images(sources, expected_images, output_path.parent)
            if expected_images and "\n|" not in text:
                raise ValueError("The saved Markdown has no table.")
        else:
            raise ValueError(f"Unknown operation: {operation}")
        required = list(markers(source))
        if expected_images:
            required.append("TableWitness")
            if result["verifiedImages"] < len(expected_images):
                raise ValueError("The converted artifact has missing image evidence.")
        for marker in required:
            if marker not in text:
                raise ValueError(f"Missing artifact witness {marker}.")
    return result


def verify_images(sources, expected, directory):
    remaining = list(expected)
    for src in sources:
        prefix = "data:image/png;base64,"
        if src.startswith(prefix):
            data = base64.b64decode(src[len(prefix):])
        else:
            path = directory / src
            if Path(src).name != src or not path.is_file():
                continue
            data = path.read_bytes()
        if data in remaining:
            remaining.remove(data)
    return len(expected) - len(remaining)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("operation")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--text-tool", default="pdftotext")
    parser.add_argument("--image-tool", default="pdfimages")
    parser.add_argument("--import-manifest", type=Path)
    parser.add_argument("--import-case")
    parser.add_argument("--import-companion", type=Path)
    arguments = parser.parse_args()
    if arguments.operation == "Import":
        if not arguments.import_manifest or not arguments.import_case:
            parser.error("Import requires a fixture manifest and case name.")
        result = verify_import(arguments.input.read_bytes(), arguments.output, arguments.import_manifest,
                               arguments.import_case, arguments.import_companion)
    else:
        result = verify(arguments.operation, arguments.input, arguments.output, arguments.text_tool, arguments.image_tool)
    print(json.dumps(result))


if __name__ == "__main__":
    main()
