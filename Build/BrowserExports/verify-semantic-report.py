"""Independent pypdf/qpdf readback of opt-in real-consumer semantic PDF artifacts.

Usage: python verify-semantic-report.py <PDF-or-directory> [--qpdf=/absolute/qpdf]
This validates the declared fixture profile, not general PDF/UA or native AT behavior.
"""
from __future__ import annotations

import argparse
from collections import Counter
import json
import math
from pathlib import Path
import re
import shutil
import subprocess

from pypdf import PdfReader
from pypdf.generic import ArrayObject, ContentStream, DictionaryObject, IndirectObject, NullObject


def require(condition, message):
    if not condition:
        raise ValueError(message)


def resolved(value):
    return value.get_object() if isinstance(value, IndirectObject) else value


def reference(value):
    if isinstance(value, IndirectObject):
        return value.idnum, value.generation
    indirect = getattr(value, "indirect_reference", None)
    return reference(indirect) if indirect is not None else None


def sequence(value):
    value = resolved(value)
    if value is None or isinstance(value, NullObject):
        return []
    return list(value) if isinstance(value, ArrayObject) else [value]


def compact(value):
    return re.sub(r"\s+", "", str(value))


def tree_entries(node, field):
    node = resolved(node)
    result = []
    entries = sequence(node.get(field))
    require(len(entries) % 2 == 0, "Odd name/number-tree entry count")
    result.extend(zip(entries[::2], entries[1::2]))
    for child in sequence(node.get("/Kids")):
        result.extend(tree_entries(child, field))
    keys = [str(key) if field == "/Names" else int(key) for key, _ in result]
    require(len(set(keys)) == len(keys), "Duplicate name/number-tree key")
    require(keys == sorted(keys), "Unsorted name/number-tree keys")
    return result


def original_bytes(value):
    if isinstance(value, bytes):
        return value
    # pypdf's method may reconstruct a UTF-16 BOM after text autodetection. The
    # public property preserves the actual bytes parsed from a content operand.
    return value.original_bytes


def unicode_map(font):
    font = resolved(font)
    if "/ToUnicode" not in font:
        return None
    cmap = resolved(font["/ToUnicode"]).get_data().decode("ascii")
    result = {}
    for block in re.findall(r"beginbfchar(.*?)endbfchar", cmap, re.S):
        for source, target in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            result[bytes.fromhex(source)] = bytes.fromhex(target).decode("utf-16-be")
    for block in re.findall(r"beginbfrange(.*?)endbfrange", cmap, re.S):
        for match in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<([0-9A-Fa-f]+)>|\[([^]]+)\])", block):
            first, last = int(match[1], 16), int(match[2], 16)
            width = len(match[1]) // 2
            if match[4]:
                start = int(match[4], 16)
                destination_width = len(match[4]) // 2
                targets = [(start + index).to_bytes(destination_width, "big").decode("utf-16-be") for index in range(last - first + 1)]
            else:
                targets = [bytes.fromhex(value).decode("utf-16-be") for value in re.findall(r"<([0-9A-Fa-f]+)>", match[5])]
            require(len(targets) == last - first + 1, "Incomplete Unicode range")
            result.update({(first + index).to_bytes(width, "big"): text for index, text in enumerate(targets)})
    require(result, "Embedded font has no supported ToUnicode entries")
    return result


def decode_text(value, mapping):
    data = original_bytes(value)
    if mapping is None:
        return data.decode("cp1252")
    widths = sorted({len(key) for key in mapping}, reverse=True)
    result, offset = [], 0
    while offset < len(data):
        found = next((data[offset:offset + width] for width in widths if data[offset:offset + width] in mapping), None)
        require(found is not None, "Text code missing from independent Unicode map")
        result.append(mapping[found]); offset += len(found)
    return "".join(result)


def page_marks(page, reader):
    fonts = {str(name): unicode_map(font) for name, font in resolved(page["/Resources"]).get("/Font", {}).items()}
    current_font, states, stack, marks = None, [], [], {}
    contents = ContentStream(page.get_contents(), reader)
    for operands, operator in contents.operations:
        if operator == b"q":
            states.append(current_font)
        elif operator == b"Q":
            current_font = states.pop()
        elif operator == b"Tf":
            current_font = str(operands[0])
        elif operator in (b"BMC", b"BDC"):
            attributes = resolved(operands[1]) if operator == b"BDC" else {}
            frame = {"artifact": str(operands[0]) == "/Artifact", "mcid": attributes.get("/MCID"), "actual": attributes.get("/ActualText"), "text": []}
            stack.append(frame)
            if frame["mcid"] is not None:
                require(int(frame["mcid"]) not in marks, "A physical MCID was used more than once")
        elif operator == b"EMC":
            require(stack, "Unbalanced marked-content end")
            frame = stack.pop()
            if frame["mcid"] is not None:
                marks[int(frame["mcid"])] = str(frame["actual"]) if frame["actual"] is not None else "".join(frame["text"])
        elif operator in (b"Tj", b"TJ", b"'", b'"'):
            values = operands[0] if operator == b"TJ" else [operands[-1]]
            text = "".join(decode_text(value, fonts[current_font]) for value in values if isinstance(value, (str, bytes)))
            if any(frame["artifact"] for frame in stack):
                continue
            owner = next((frame for frame in reversed(stack) if frame["mcid"] is not None), None)
            if owner is not None:
                owner["text"].append(text)
    require(not stack and not states, "Unbalanced content/graphics state")
    return marks


def attributes(element):
    result = {}
    for value in sequence(element.get("/A")):
        result.update(resolved(value))
    return result


def physical_group_contexts(page_texts, paths, continuation_suffix):
    """Recognize only the fixture's exact authored context lines, without inferring product roles."""
    variants = {}
    for path in paths:
        for level, label in enumerate(path):
            identity = compact(label)
            for text in (label, label + " " + continuation_suffix) if continuation_suffix else (label,):
                previous = variants.setdefault(compact(text), (identity, level))
                require(previous == (identity, level), "Physical-context fixture labels are ambiguous across levels")
    result = []
    for text in page_texts:
        runs, current = [], []
        for line in text.splitlines():
            normalized = compact(line)
            if normalized in variants:
                current.append(variants[normalized])
            elif normalized and current:
                runs.append(current); current = []
        if current:
            runs.append(current)
        result.append(runs)
    return result


def has_physical_context(runs, path):
    expected = tuple(compact(label) for label in path)
    for run in runs:
        for first in range(len(run) - len(expected) + 1):
            if tuple(label for label, _ in run[first:first + len(expected)]) != expected:
                continue
            following = first + len(expected)
            # A stale descendant caption contains its parent's prefix but is not that parent's
            # summary context. Another root/sibling context in the same physical page is allowed.
            if following < len(run) and run[following][1] >= len(expected):
                continue
            return True
    return not expected


def verify(path, qpdf):
    contract = json.loads(path.with_suffix(path.suffix + ".json").read_text())
    labels = contract.get("labels", {"groupHeader": "Group", "subtotal": "Subtotal", "grandTotal": "Grand total"})
    process = subprocess.run([qpdf, "--check", str(path)], capture_output=True, text=True, timeout=60)
    require(process.returncode == 0, "qpdf rejected " + str(path) + ": " + process.stderr)
    reader = PdfReader(path, strict=True)
    catalog = reader.trailer["/Root"]
    require(bool(catalog["/MarkInfo"]["/Marked"]) and str(catalog["/Lang"]) == contract["language"], "Missing marked/language catalog contract")
    root = catalog["/StructTreeRoot"]
    parents = dict((int(key), value) for key, value in tree_entries(root["/ParentTree"], "/Nums"))
    identifiers = dict((str(key), value) for key, value in tree_entries(root["/IDTree"], "/Names")) if "/IDTree" in root else {}
    pages = {reference(page): (index, page) for index, page in enumerate(reader.pages)}
    marks = {key: page_marks(page, reader) for key, (_, page) in pages.items()}
    consumed, elements, annotations = set(), {}, set()

    def visit(value, parent, ancestry, inherited_page=None):
        obj = resolved(value)
        if isinstance(obj, int):
            return mark_text(inherited_page, obj, parent)
        if not isinstance(obj, DictionaryObject):
            return ""
        page_key = reference(obj.get("/Pg")) or inherited_page
        if str(obj.get("/Type")) == "/MCR":
            return mark_text(page_key, int(obj["/MCID"]), parent)
        if str(obj.get("/Type")) == "/OBJR":
            annotation = resolved(obj["/Obj"]); annotation_id = reference(obj["/Obj"])
            require(annotation_id not in annotations, "Duplicate logical annotation object")
            require(reference(parents[int(annotation["/StructParent"])]) == parent, "Annotation parent-tree owner mismatch")
            require(reference(annotation.get("/P")) == page_key, "Annotation page binding differs")
            annotations.add(annotation_id)
            return ""
        if "/S" not in obj:
            return "".join(visit(child, parent, ancestry, page_key) for child in sequence(obj.get("/K")))
        identity = reference(value)
        require(identity is not None and identity not in elements, "Duplicate/non-indirect logical structure element")
        require(reference(obj["/P"]) == parent, "Logical parent link differs")
        entry = {"id": identity, "object": obj, "role": str(obj["/S"]), "ancestors": ancestry[:], "children": sequence(obj.get("/K")), "pages": set()}
        elements[identity] = entry
        entry["text"] = "".join(visit(child, identity, [*ancestry, identity], page_key) for child in entry["children"])
        for child in entry["children"]:
            if reference(child) in elements:
                entry["pages"].update(elements[reference(child)]["pages"])
        return entry["text"]

    def mark_text(page_key, mcid, parent):
        require(page_key in pages and mcid in marks[page_key], "Logical reference has no physical marked text")
        key = page_key, mcid
        require(key not in consumed, "One physical marked fragment has duplicate logical owners")
        consumed.add(key)
        page = pages[page_key][1]
        owners = sequence(parents[int(page["/StructParents"])])
        require(mcid < len(owners) and reference(owners[mcid]) == parent, "Page parent-tree owner mismatch")
        elements[parent]["pages"].add(page_key)
        return marks[page_key][mcid]

    logical_text = visit(root, reference(root), [])
    require(len(consumed) == sum(len(value) for value in marks.values()), "Physical marked content is missing from the logical tree")
    for name, owner in identifiers.items():
        require(reference(owner) in elements and str(resolved(owner).get("/ID")) == name, "IDTree target differs from its structure ID")
    for entry in elements.values():
        for name in sequence(attributes(entry["object"]).get("/Headers")):
            require(str(name) in identifiers, "Unresolved header association")
            target = elements[reference(identifiers[str(name)])]
            require(target["role"] == "/TH", "Cell header association targets a non-heading")
    by_role = Counter(entry["role"] for entry in elements.values())
    rows = [entry for entry in elements.values() if entry["role"] == "/TR"]
    bodies = [entry for entry in rows if re.search(r"S\d{4}", compact(entry["text"]))]
    expected_ids = [row[0] for row in contract["rows"]]

    def owning_table(entry):
        return next((identity for identity in reversed(entry["ancestors"]) if elements[identity]["role"] == "/Table"), None)

    def expected_columns(first, span=1):
        if contract["grouped"]:
            names = [labels["groupHeader"], *[column["header"] for column in contract["columns"]]]
            return [compact(value) for value in names[first:first + span]]
        result = []
        for row in contract["headers"]:
            for column, anchor in enumerate(row):
                if anchor and column < first + span and first < column + anchor.get("columnSpan", 1) and compact(anchor["value"]):
                    result.append(compact(anchor["value"]))
        return result

    def column_associations(entry):
        targets = [elements[reference(identifiers[str(name)])] for name in sequence(attributes(entry["object"]).get("/Headers"))]
        require(all(owning_table(target) == owning_table(entry) for target in targets), "Cell refers to another logical table's headings")
        return [compact(target["text"]) for target in targets if str(attributes(target["object"]).get("/Scope")) == "/Column"]

    require([re.findall(r"S\d{4}", compact(row["text"])) for row in bodies] == [[identifier] for identifier in expected_ids], "Source IDs are missing, duplicated, or in visual-band order in the logical tree")
    for row, expected in zip(bodies, contract["rows"]):
        cells = [elements[reference(child)] for child in row["children"] if reference(child) in elements]
        actual = [compact(cell["text"]) for cell in cells]
        require(actual == ([""] if contract["grouped"] else []) + [compact(value) for value in expected], "Logical source cell values/order differ for " + expected[0])
        require(all(sequence(attributes(cell["object"]).get("/Headers")) for cell in cells), "A body cell lacks its column-header relationships")
        for column, cell in enumerate(cells):
            require(column_associations(cell) == expected_columns(column), "Body cell column-header identity/order differs for " + expected[0])
    require(compact(contract["expectedFragmentToken"]) in compact(logical_text), "Last overheight-cell fragment is missing")
    page_texts = [page.extract_text() or "" for page in reader.pages]
    physical_pages = [compact(text) for text in page_texts]
    physical = "".join(page_texts)
    for forbidden in contract["forbidden"]:
        require(compact(forbidden) not in compact(logical_text) and compact(forbidden) not in compact(physical), "Excluded host data leaked")
    require(all(compact(logical_text).count(identifier) == 1 for identifier in expected_ids), "A repeated visual ID appears more than once in logical reading")
    require(all(physical.count(identifier) >= 2 for identifier in expected_ids), "Fixture did not paint repeated identifier bands")

    def anchors(matrix):
        return [cell for row in matrix or [] for cell in row if cell is not None]

    for anchor in anchors(contract.get("headers")) + anchors(contract.get("footers")):
        candidates = [entry for entry in elements.values() if entry["role"] in ("/TH", "/TD") and compact(entry["text"]) == compact(anchor["value"])]
        require(len(candidates) == 1, "Span anchor was duplicated/lost in logical reading: " + str(anchor["value"]))
        attr = attributes(candidates[0]["object"])
        require(int(attr.get("/ColSpan", 1)) == anchor.get("columnSpan", 1) and int(attr.get("/RowSpan", 1)) == anchor.get("rowSpan", 1), "Semantic span differs from captured host rectangle")
    for row in contract.get("footers") or []:
        for column, anchor in enumerate(row):
            if anchor is None:
                continue
            candidate = next(entry for entry in elements.values() if entry["role"] in ("/TH", "/TD") and compact(entry["text"]) == compact(anchor["value"]))
            require(column_associations(candidate) == expected_columns(column, anchor.get("columnSpan", 1)), "Footer span lost its exact column-header associations")
    checked_contexts = set()
    if contract["grouped"]:
        paths = list(dict.fromkeys(tuple(path[:level + 1]) for path in contract["groupPaths"] for level in range(len(path))))
        contexts = physical_group_contexts(page_texts, contract["groupPaths"], contract.get("continuationSuffix", ""))
        headings = [entry for entry in elements.values() if entry["role"] in ("/H2", "/H3", "/H4", "/H5")]
        require([(int(entry["role"][2:]) - 2, compact(entry["text"])) for entry in headings] == [(len(path) - 1, compact(path[-1])) for path in paths], "Hierarchical logical headings differ")
        def group_ancestry(entry):
            groups = []
            for ancestor in entry["ancestors"]:
                node = elements[ancestor]
                if node["role"] != "/Sect":
                    continue
                direct = [elements[reference(child)] for child in node["children"] if reference(child) in elements]
                groups.extend(child["text"] for child in direct if child["role"] in ("/H2", "/H3", "/H4", "/H5"))
            return tuple(compact(value) for value in groups)

        for row, expected_path, identifier in zip(bodies, contract["groupPaths"], expected_ids):
            require(group_ancestry(row) == tuple(compact(value) for value in expected_path), "A record is attached to the wrong ancestor section")
            detail_pages = [index for index, text in enumerate(physical_pages) if compact(identifier) in text]
            require(detail_pages, "A source ID has no physical detail page")
            for index in detail_pages:
                require(has_physical_context(contexts[index], expected_path),
                        "Physical detail page " + str(index + 1) + " lacks the full active ancestor context for " + identifier)
                checked_contexts.add((index, tuple(compact(value) for value in expected_path)))
        summaries = [entry for entry in elements.values() if entry["role"] == "/TH" and str(attributes(entry["object"]).get("/Scope")) == "/Row"]
        require(len(summaries) == len(paths) + 1, "Subtotal/grand-total logical row headings differ")
        source_rows = contract.get("rawRows", contract["rows"])
        require(len(source_rows) == len(contract["groupPaths"]), "Summary oracle source/path counts differ")
        require(all(len(row) == len(contract["columns"]) for row in source_rows), "Raw summary source shape differs")
        require([str(row[0]) for row in source_rows] == expected_ids, "Raw summary source record order differs")
        aggregates = contract.get("aggregates", {"c6": "sum", "c8": "max"})
        require(aggregates and all(operation in ("sum", "count", "average", "min", "max") for operation in aggregates.values()), "Missing/unsupported fixture aggregate contract")
        metric_positions = {column["key"]: index for index, column in enumerate(contract["columns"]) if column["key"] in aggregates}
        require(set(metric_positions) == set(aggregates), "A declared aggregate does not identify one source column")

        def summary_values(indices, label):
            values = [""] * len(contract["columns"])
            for key, position in metric_positions.items():
                raw_values = [source_rows[index][position] for index in indices]
                if "rawRows" in contract:
                    require(all(value is None or type(value) in (int, float) for value in raw_values), "Typed summary source lost a numeric metric")
                numbers = [float(value) for value in raw_values if value is not None and value != ""]
                require(all(math.isfinite(value) for value in numbers), "Summary oracle received a non-finite metric")
                operation = aggregates[key]
                value = math.fsum(numbers) if operation == "sum" else len(numbers) if operation == "count" else None
                if numbers and operation in ("average", "min", "max"):
                    value = math.fsum(numbers) / len(numbers) if operation == "average" else min(numbers) if operation == "min" else max(numbers)
                values[position] = "" if value is None else format(value, ".15g")
            return [compact(label), *values]

        label_overrides = {tuple(compact(value) for value in entry["path"]): entry["text"] for entry in contract.get("summaryLabels", [])}
        if "summaryLabels" in contract:
            require(len(label_overrides) == len(contract["summaryLabels"]) and set(label_overrides) == {tuple(compact(value) for value in path) for path in paths} | {()},
                    "Authored summary-label paths do not match the source hierarchy")
        expected_summaries = {}
        for group_path in paths:
            indices = [index for index, source_path in enumerate(contract["groupPaths"]) if tuple(source_path[:len(group_path)]) == group_path]
            identity = tuple(compact(value) for value in group_path)
            label = label_overrides.get(identity, labels["subtotal"] + " " + group_path[-1] + " (" + str(len(indices)) + ")")
            expected_summaries[identity] = summary_values(indices, label)
        expected_summaries[()] = summary_values(range(len(source_rows)), label_overrides.get((), labels["grandTotal"] + " (" + str(len(source_rows)) + ")"))
        summary_label_counts = Counter(values[0] for values in expected_summaries.values())
        seen_summaries = set()
        for summary in summaries:
            row = elements[reference(summary["object"]["/P"])]; identifier = str(summary["object"]["/ID"])
            owner_path = group_ancestry(row)
            require(owner_path in expected_summaries and owner_path not in seen_summaries, "Summary was duplicated or attached to the wrong ancestor section")
            seen_summaries.add(owner_path)
            cells = [elements[reference(child)] for child in row["children"] if reference(child) in elements]
            require([compact(cell["text"]) for cell in cells] == expected_summaries[owner_path], "Summary label/count/independently computed numeric totals differ")
            if owner_path:
                summary_pages = {pages[key][0] for key in row["pages"]}
                label = expected_summaries[owner_path][0]
                if summary_label_counts[label] == 1:
                    summary_pages.update(index for index, text in enumerate(physical_pages) if label in text)
                require(summary_pages, "A subtotal has no physical summary page")
                for index in summary_pages:
                    require(has_physical_context(contexts[index], owner_path),
                            "Physical subtotal page " + str(index + 1) + " lacks its own active ancestor context: " + "/".join(owner_path))
                    checked_contexts.add((index, owner_path))
            for column, child in enumerate(row["children"]):
                entry = elements.get(reference(child))
                if entry is not None:
                    require(column_associations(entry) == expected_columns(column), "Summary column-header association differs")
                    if entry is not summary:
                        require(identifier in [str(value) for value in sequence(attributes(entry["object"]).get("/Headers"))], "Summary value lost its explicit row-header association")
        require(seen_summaries == set(expected_summaries), "One or more source-derived summaries are missing")
        def outline_titles(items):
            result = []
            for item in items:
                result.extend(outline_titles(item) if isinstance(item, list) else [str(item["/Title"])])
            return result
        titles = outline_titles(reader.outline)
        expected_groups = [path[-1] for path in paths]
        require([title for title in titles if title in set(expected_groups)] == expected_groups, "Native outline loses group navigation/order")
    links = [resolved(item) for page in reader.pages for item in sequence(page.get("/Annots")) if str(resolved(item).get("/Subtype")) == "/Link"]
    require(links and len(annotations) == len(links), "Native links and logical OBJR associations differ")
    for prefix in contract["links"]:
        require(any(str(link.get("/A", {}).get("/URI", "")).startswith(prefix) for link in links), "Required native link target missing")
    return {"file": str(path), "pages": len(reader.pages), "sourceRows": len(bodies), "logicalElements": len(elements), "roles": dict(by_role),
            "markedFragments": len(consumed), "headerIdentities": len(identifiers), "nativeLinks": len(links), "qpdf": "PASS",
            "sourceOrderExactlyOnce": True, "physicalRepeatedIds": True, "summaryValuesAndOwnership": bool(contract["grouped"]),
            "physicalAncestorContextPagePaths": len(checked_contexts), "physicalContextProof": "exact authored context-line presence/order, not per-row native-reader or AT acceptance",
            "typedSummarySource": bool(contract.get("rawRows")), "independentReader": "pypdf", "nativeAT": "separate acceptance"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path)
    parser.add_argument("--qpdf", default=shutil.which("qpdf"))
    args = parser.parse_args()
    require(args.qpdf, "Supply installed qpdf with --qpdf")
    files = [args.path] if args.path.is_file() else sorted(args.path.rglob("*.pdf"))
    require(files, "No semantic report artifacts found")
    reports = [verify(path, args.qpdf) for path in files]
    output = args.path.parent if args.path.is_file() else args.path
    (output / "semantic-independent-report.json").write_text(json.dumps({"reports": reports}, indent=2) + "\n")
    print(json.dumps({"files": len(reports), "pages": sum(report["pages"] for report in reports), "result": "PASS"}))


if __name__ == "__main__":
    main()
