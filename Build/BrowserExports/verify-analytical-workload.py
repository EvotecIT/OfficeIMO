"""Independent stdlib/Poppler readback of the analytical composition workload.

Usage: python verify-analytical-workload.py output count xlsx|pdf [pdftotext]
The expected records and hierarchy are derived here without loading the JS exporter.
"""
import json
import math
import re
import subprocess
import sys
import zipfile
from datetime import datetime
import xml.etree.ElementTree as ET

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "h": "http://www.w3.org/1999/xhtml"}
TAG = "{" + NS["s"] + "}"
HEADERS = ["Group", "Service", "Site", "Status", "Latency (ms)", "Availability", "Checked (UTC)"]
SCOPE = "Illustrative selected inventory. Grouping fields are metadata; detail values remain typed."


def require(condition, message):
    if not condition:
        raise ValueError(message)


def values(index, count):
    leaf = index // max(1, math.ceil(count / 24))
    return ["WEB-" + str(index + 1).zfill(6), ["Warsaw", "Łódź", "Praha", "Berlin"][leaf % 4],
            "Offline" if index % 37 == 0 else "Attention" if index % 11 == 0 else "Healthy",
            None if index % 17 == 0 else index % 97 / 10 + 4.75,
            None if index % 23 == 0 else .995 - index % 17 / 10000,
            datetime(2026, 10, 21, 8, index % 60)]


def metrics(first, end, count):
    latency, availability = [], []
    for index in range(first, end):
        row = values(index, count)
        if row[3] is not None:
            latency.append(row[3])
        if row[4] is not None:
            availability.append(row[4])
    return [math.fsum(items) / len(items) if items else None for items in [latency, availability]]


def hierarchy(count):
    roots, lookup = [], {}
    leaf_rows = max(1, math.ceil(count / 24))
    for leaf, first in enumerate(range(0, count, leaf_rows)):
        path = ["Portfolio " + str(leaf // 12), ["Production", "Staging", "Development"][(leaf // 4) % 3], "Owner " + str(leaf % 4)]
        siblings = roots
        for level, name in enumerate(path):
            identity = tuple(path[:level + 1])
            if identity not in lookup:
                node = {"label": name, "path": identity, "level": level, "first": first, "end": min(count, first + leaf_rows), "children": []}
                lookup[identity] = node
                siblings.append(node)
            node = lookup[identity]
            node["end"] = min(count, first + leaf_rows)
            siblings = node["children"]
    return roots, list(lookup.values())


def physical_layout(roots):
    events, details, row = [], [], 3  # Title and leaf headings occupy rows 1 and 2.

    def visit(node):
        nonlocal row
        events.append((row, "group", node)); row += 1
        node["firstDetail"] = row + (2 - node["level"])
        if node["children"]:
            for child in node["children"]:
                visit(child)
            node["lastDetail"] = node["children"][-1]["lastDetail"]
        else:
            first_row = row
            events.append((row, "details", node))
            row += node["end"] - node["first"]
            node["lastDetail"] = row - 1
            details.append((first_row, row - 1))
        node["summaryRow"] = row
        events.append((row, "summary", node)); row += 1

    for root in roots:
        visit(root)
    return events, details, row


def cell_text(cell):
    return "".join(cell.find("s:is", NS).itertext()) if cell.get("t") == "inlineStr" else cell.findtext("s:v", "", NS)


def numeric(cell):
    require(cell.get("t") in (None, "n"), "Numeric/date/formula cache has a text cell type")
    number = float(cell.findtext("s:v", "", NS))
    require(math.isfinite(number), "A numeric value is not finite")
    return number


def blank(cell):
    return cell.find("s:v", NS) is None and cell.find("s:is", NS) is None and cell.find("s:f", NS) is None


def xlsx(source, count):
    roots, nodes = hierarchy(count)
    events, ranges, grand_row = physical_layout(roots)
    event_index, detail_index, actual_rows = 0, None, 0
    cache_count = 0
    with zipfile.ZipFile(source) as package:
        require(package.testzip() is None, "ZIP CRC differs")
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        sheets = workbook.findall("s:sheets/s:sheet", NS)
        require([item.get("name") for item in sheets] == ["Report context", "Grouped inventory", "Group overview"], "Composed workbook must have exactly its three visible sheets")
        require(all(item.get("state", "visible") == "visible" for item in sheets), "A hidden duplicate source sheet was created")
        calc = workbook.find("s:calcPr", NS)
        if count:
            require(calc is not None and calc.get("calcMode") == "auto" and calc.get("forceFullCalc") == "1", "Editable group formulas need automatic recalculation")
        context = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
        context_rows = context.findall("s:sheetData/s:row", NS)
        require([cell_text(item.find("s:c", NS)) for item in context_rows] == ["Report context", "Selected inventory analysis", SCOPE], "Pre-detail context content differs")

        def summary(cells, first, last, expected, label):
            nonlocal cache_count
            require(cell_text(cells[0]) == label, "Group/grand summary label or selected-row count differs")
            require(all(blank(cells[index]) for index in [1, 2, 3, 6]), "Summary overwrote an untotaled source column")
            for index, column, cached in zip([4, 5], "EF", expected):
                cell = cells[index]
                if not count:
                    require(blank(cell), "Empty average must be blank")
                    continue
                span = column + str(first) + ":" + column + str(last)
                expected_formula = 'IF(SUBTOTAL(2,' + span + ')=0,"",SUBTOTAL(1,' + span + '))'
                require(cell.findtext("s:f", "", NS) == expected_formula, "A summary lost its finite native SUBTOTAL reference")
                if cached is None:
                    require(cell.get("t") == "str" and cell.findtext("s:v", "", NS) == "", "An all-null average cache became numeric")
                else:
                    require(abs(numeric(cell) - cached) < 1e-9, "A summary cache differs from its raw source records")
                cache_count += 1

        with package.open("xl/worksheets/sheet2.xml") as stream:
            parser = ET.iterparse(stream, events=("end",))
            for _, element in parser:
                if element.tag != TAG + "row":
                    continue
                actual_rows += 1
                require(int(element.get("r")) == actual_rows, "Physical worksheet row order differs")
                cells = element.findall("s:c", NS)
                require(len(cells) == 7 and [cell.get("r") for cell in cells] == [column + str(actual_rows) for column in "ABCDEFG"], "A row lost, duplicated or moved a visible source cell")
                level = 0
                if actual_rows == 1:
                    require(cell_text(cells[0]) == "Selected inventory analysis", "Analysis title differs")
                elif actual_rows == 2:
                    require([cell_text(cell) for cell in cells] == HEADERS, "Visible schema gained a hidden grouping field or changed order")
                elif actual_rows == grand_row:
                    summary(cells, ranges[0][0] if ranges else 0, ranges[-1][1] if ranges else 0, metrics(0, count, count), "All selected records · " + format(count, ","))
                else:
                    expected_row, kind, node = events[event_index]
                    level = node["level"]
                    if kind == "details":
                        if detail_index is None:
                            detail_index = node["first"]
                        level = 3
                        require(actual_rows == expected_row + detail_index - node["first"], "Detail coordinate differs")
                        require(blank(cells[0]), "The generated label column overwrote a typed source cell")
                        for cell, value in zip(cells[1:], values(detail_index, count)):
                            require(cell.find("s:f", NS) is None, "A literal captured source value became a formula")
                            if isinstance(value, str):
                                require(cell.get("t") == "inlineStr" and cell_text(cell) == value, "Source identity/order/text differs at " + str(detail_index))
                            elif value is None:
                                require(blank(cell), "A null metric acquired a stored value at " + str(detail_index))
                            else:
                                expected = (value - datetime(1899, 12, 30)).total_seconds() / 86400 if isinstance(value, datetime) else value
                                require(abs(numeric(cell) - expected) < 1e-9, "Typed number/date differs at " + str(detail_index))
                        detail_index += 1
                        if detail_index == node["end"]:
                            event_index += 1; detail_index = None
                    else:
                        require(actual_rows == expected_row, "Generated group coordinate differs")
                        if kind == "group":
                            require(cell_text(cells[0]) == node["label"] and all(blank(cell) for cell in cells[1:]), "Group header identity differs")
                        else:
                            summary(cells, node["firstDetail"], node["lastDetail"], metrics(node["first"], node["end"], count),
                                    "Group total · " + node["label"] + " · " + format(node["end"] - node["first"], ","))
                        event_index += 1
                require(int(element.get("outlineLevel", "0")) == level and element.get("hidden", "0") == "0", "Native three-level outline ownership differs")
                element.clear()
        require(actual_rows == grand_row and event_index == len(events), "Detail/group/grand physical row count differs")
        worksheet = parser.root
        require(worksheet.find("s:sheetFormatPr", NS).get("outlineLevelRow") == "3", "Worksheet maximum outline depth differs")
        require(worksheet.find("s:autoFilter", NS) is None and worksheet.find("s:tableParts", NS) is None, "Native filtering must not include generated summary rows")
        pane = worksheet.find("s:sheetViews/s:sheetView/s:pane", NS)
        require(pane is not None and pane.get("xSplit") == "1" and pane.get("ySplit") == "2", "Native frozen label/header coordinates differ")
        selected = " ".join("D" + str(first) + ":D" + str(last) for first, last in ranges)
        formats = worksheet.findall("s:conditionalFormatting", NS)
        validations = worksheet.findall("s:dataValidations/s:dataValidation", NS)
        require(len(formats) == (1 if count else 0) and len(validations) == (1 if count else 0), "Detail-only native editing-rule count differs")
        if count:
            rule = formats[0].find("s:cfRule", NS)
            require(formats[0].get("sqref") == selected and rule.get("type") == "cellIs" and rule.get("operator") == "equal" and rule.findtext("s:formula", "", NS) == '"Offline"', "Conditional formatting targets a generated row or changed its status comparison")
            require(validations[0].get("sqref") == selected and validations[0].get("allowBlank") == "0" and validations[0].findtext("s:formula1", "", NS) == '"Healthy,Attention,Offline"', "Dropdown scope or choices differ")
        overview = ET.fromstring(package.read("xl/worksheets/sheet3.xml"))
        overview_rows = overview.findall("s:sheetData/s:row", NS)
        require(len(overview_rows) == len(roots) + 1 and [cell_text(cell) for cell in overview_rows[0].findall("s:c", NS)] == ["Group", "Records"], "Post-detail overview schema/row count differs")
        for row, node in zip(overview_rows[1:], roots):
            cells = row.findall("s:c", NS)
            require(cell_text(cells[0]) == node["label"] and numeric(cells[1]) == node["end"] - node["first"], "Post-detail overview did not reuse the one-pass group summary")
        charts = [name for name in package.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", name)]
        require(len(charts) == (1 if roots else 0), "Native overview chart count differs")
        if roots:
            chart = ET.fromstring(package.read(charts[0]))
            require([item.text for item in chart.findall(".//c:f", NS)] == ["'Group overview'!$B$1", "'Group overview'!$A$2:$A$" + str(len(roots) + 1), "'Group overview'!$B$2:$B$" + str(len(roots) + 1)], "Chart references escaped the finite visible overview")
            require([item.text for item in chart.findall(".//c:cat/c:strRef/c:strCache/c:pt/c:v", NS)] == [node["label"] for node in roots], "Chart category cache differs")
            require([float(item.text) for item in chart.findall(".//c:numCache/c:pt/c:v", NS)] == [node["end"] - node["first"] for node in roots], "Chart numeric cache differs")
    return {"rows": count, "sourceColumns": 6, "columns": 7, "everyCell": True, "zipCrc": True, "typedNumbersNullsUtcDatesAndIds": True,
            "groups": len(nodes), "levels": 3, "numericFormulaCaches": cache_count, "finiteSubtotalReferences": True,
            "detailOnlyEditingRules": True, "detailRuns": len(ranges), "visibleSheets": 3, "postDetailOverviewRows": len(roots), "boundChartPoints": len(roots)}


def pdf(source, count, pdftotext):
    extracted = subprocess.run([pdftotext, "-bbox-layout", "-enc", "UTF-8", source, "-"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf8").stdout
    document = ET.fromstring(extracted)
    roots, nodes = hierarchy(count)
    events, _, _ = physical_layout(roots)
    summaries = [node for _, kind, node in events if kind == "summary"]
    body_index, summary_index, grand_count, continuations = 0, 0, 0, 0
    header_positions = None
    all_text = []

    def number_text(actual, expected, message):
        require(actual == "" if expected is None else bool(re.fullmatch(r"\d+(?:\.\d{1,3})?", actual)) and abs(float(actual) - expected) <= .000500001, message)

    pages = document.findall(".//h:page", NS)
    for page_index, page in enumerate(pages):
        words = sorted([(float(word.get("yMin")), float(word.get("xMin")), word.text or "") for word in page.findall(".//h:word", NS)])
        lines = []
        for y, x, text in words:
            if not lines or abs(lines[-1][0] - y) > .1:
                lines.append((y, []))
            lines[-1][1].append((x, text))
        lines = [(y, sorted(items)) for y, items in lines]
        texts = [" ".join(text for _, text in items) for _, items in lines]
        all_text.extend(texts)
        local_header = None
        for position, (_, items) in enumerate(lines):
            text = texts[position]
            by_text = {text: x for x, text in items}
            if all(name in by_text for name in ["Service", "Site", "Status", "Latency", "Availability", "Checked"]):
                header_positions = [by_text[name] for name in ["Service", "Site", "Status", "Latency", "Availability", "Checked"]]
                local_header = (position, header_positions)
                continue
            ids = [value for _, value in items if re.fullmatch(r"WEB-\d{6}", value)]
            if ids:
                require(len(ids) == 1 and local_header is not None, "A body page lost its repeated visible schema")
                require(float(page.get("width")) > float(page.get("height")), "Analytical body pages must use their landscape section")
                expected = values(body_index, count)
                require(ids[0] == expected[0], "PDF lost, duplicated or reordered a selected record")
                boundaries = local_header[1]
                cells = [" ".join(value for x, value in items if x >= first - .5 and (column == 5 or x < boundaries[column + 1] - .5)) for column, first in enumerate(boundaries)]
                require(cells[:3] == expected[:3] and cells[5] == expected[5].strftime("%Y-%m-%d %H:%M"), "PDF identity/site/status/date differs at " + str(body_index))
                number_text(cells[3], expected[3], "PDF latency/null/display precision differs at " + str(body_index))
                number_text(cells[4], expected[4], "PDF availability/null/display precision differs at " + str(body_index))
                leaf = body_index // max(1, math.ceil(count / 24))
                path = ["Portfolio " + str(leaf // 12), ["Production", "Staging", "Development"][(leaf // 4) % 3], "Owner " + str(leaf % 4)]
                captions = texts[max(0, local_header[0] - 3):local_header[0]]
                require(captions == path or captions == [name + " (continued)" for name in path], "A PDF detail fragment lost its full three-level active group context")
                body_index += 1
            elif text.startswith("Group total · ") or text.startswith("All selected records ·"):
                require(header_positions is not None or count == 0, "A PDF summary has no preceding analytical table")
                grand = text.startswith("All selected records ·")
                if grand:
                    grand_count += 1; node = {"first": 0, "end": count}
                    expected_label = "All selected records · " + format(count, ",")
                else:
                    require(summary_index < len(summaries), "PDF duplicated a group closing summary")
                    node = summaries[summary_index]; summary_index += 1
                    expected_label = "Group total · " + node["label"] + " · " + format(node["end"] - node["first"], ",")
                # Group labels may wrap within their leading cell; metrics stay in their typed source columns.
                label_end = header_positions[0] if header_positions else float("inf")
                actual_label = " ".join(value for x, value in items if x < label_end - .5)
                following = position + 1
                while actual_label != expected_label and following < len(lines) and expected_label.startswith(actual_label + " "):
                    additional = " ".join(value for x, value in lines[following][1] if x < label_end - .5)
                    if not additional:
                        break
                    actual_label += " " + additional; following += 1
                require(actual_label == expected_label, "PDF group/grand label or selected-record count differs: " + actual_label)
                if header_positions:
                    for column, expected_metric in zip([3, 4], metrics(node["first"], node["end"], count)):
                        actual = " ".join(value for x, value in items if x >= header_positions[column] - .5 and x < header_positions[column + 1] - .5)
                        number_text(actual, expected_metric, "PDF group/grand numeric aggregate differs")
            if text.endswith(" (continued)"):
                continuations += 1
    require(body_index == count and summary_index == len(nodes) and grand_count == 1, "PDF body, group subtotal or grand-total count differs")
    require("Selected inventory analysis" in all_text and "Grouped inventory" in all_text and "Next actions" in all_text, "PDF lost a composed section")
    require(any("Assign owners to the findings" in text for text in all_text), "PDF lost its closing advice")
    require(format(count, ",") + " selected records were consumed once." in all_text, "PDF closing count differs from the selected capture")
    if count >= 1000:
        require(continuations >= 3, "The multipage workload did not exercise full-path continuation captions")
    return {"rows": count, "sourceColumns": 6, "columns": 7, "everyCell": True, "typedMetricTextNullsUtcDatesAndIds": True,
            "groups": len(nodes), "levels": 3, "groupSummaries": summary_index, "grandSummaries": grand_count, "pages": len(pages),
            "continuationCaptions": continuations, "fullContextOnEveryBodyFragment": True, "composedClosingSections": True}


if __name__ == "__main__":
    require(len(sys.argv) >= 4, "Expected output, count, format and optional pdftotext path")
    output, row_count, output_format = sys.argv[1], int(sys.argv[2]), sys.argv[3]
    require(0 <= row_count <= 100000, "Workload count must be within 0 through 100,000")
    require(output_format in ("xlsx", "pdf"), "Expected xlsx or pdf")
    proof = xlsx(output, row_count) if output_format == "xlsx" else pdf(output, row_count, sys.argv[4] if len(sys.argv) > 4 else "pdftotext")
    print(json.dumps(proof, ensure_ascii=False))
