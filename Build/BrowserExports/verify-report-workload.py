"""Independent ZIP/XML readback for the opt-in deterministic workload."""
import json
import sys
import zipfile
import math
import re
import subprocess
from datetime import datetime
import xml.etree.ElementTree as ET

source, count = sys.argv[1], int(sys.argv[2])
mode = sys.argv[3] if len(sys.argv) > 3 else "inventory"
group_rows = int(sys.argv[4]) if mode in ("grouped-analysis", "grouped-pdf") else 0
ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def text(cell):
    return "".join(cell.find("s:is", ns).itertext()) if cell.get("t") == "inlineStr" else cell.findtext("s:v", "", ns)


def number(cell):
    require(cell.get("t") in (None, "n"), "Expected a numeric/date or numeric formula-cache cell")
    value = float(cell.findtext("s:v", "", ns))
    require(math.isfinite(value), "Numeric cell is not finite")
    return value


def grouped_values(index):
    group = index // group_rows
    empty = group == 0
    return ["SRV-" + str(index + 1).zfill(6), "Site" + str(group).zfill(2),
            None if empty or index % 17 == 0 else index % 97 / 10 + 4.75,
            None if empty or index % 23 == 0 else (9950 - index % 17) / 10000,
            None if empty or index % 29 == 0 else index % 31,
            datetime(2026, 10, 20, 0, index % 60)]


def grouped_totals(first, last):
    metrics = [[], [], []]
    for index in range(first, last):
        values = grouped_values(index)
        for column in range(3):
            if values[column + 2] is not None:
                metrics[column].append(values[column + 2])
    return [math.fsum(metrics[0]) / len(metrics[0]) if metrics[0] else None,
            min(metrics[1]) if metrics[1] else None, sum(metrics[2])]


def grouped_analysis(package):
    groups = (count + group_rows - 1) // group_rows
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    sheets = workbook.findall("s:sheets/s:sheet", ns)
    require(len(sheets) == 1 and sheets[0].get("name") == "Analysis" and sheets[0].get("state", "visible") == "visible", "Expected one visible analytical sheet")
    require(workbook.find("s:calcPr", ns).get("calcMode") == "auto", "Editable formulas need native recalculation")
    worksheet = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
    rows = worksheet.findall("s:sheetData/s:row", ns)
    require(len(rows) == count + groups * 2 + 2, "Grouped detail/heading/subtotal/grand row count differs")
    require([text(cell) for cell in rows[0].findall("s:c", ns)] == ["Group", "Service", "Site", "Latency (ms)", "Availability", "Requests", "Checked (UTC)"], "Grouped headings differ")
    require(worksheet.find("s:sheetFormatPr", ns).get("outlineLevelRow") == "1", "Native outline metadata differs")
    coordinates = {int(element.get("r")): element for element in rows}
    cache_count = 0

    def summary(element, first, last, expected, label):
        nonlocal cache_count
        cells = element.findall("s:c", ns)
        require(len(cells) == 7 and text(cells[0]) == label and element.get("outlineLevel", "0") == "0", "Summary label/outline differs")
        for column, operation, cached in zip("DEF", (1, 5, 9), expected):
            cell = cells[ord(column) - ord("A")]
            span = column + str(first) + ":" + column + str(last)
            expression = "SUBTOTAL(" + str(operation) + "," + span + ")"
            if operation != 9:
                expression = 'IF(SUBTOTAL(2,' + span + ')=0,"",' + expression + ')'
            require(cell.findtext("s:f", "", ns) == expression, "Finite summary formula reference differs")
            if cached is None:
                require(cell.get("t") == "str" and text(cell) == "", "Empty average/minimum cache must stay blank")
            else:
                require(abs(number(cell) - cached) < 1e-8, "Summary numeric cache differs")
            cache_count += 1

    last_detail = 0
    for group in range(groups):
        first_index = group * group_rows
        end_index = min(count, first_index + group_rows)
        first_detail = 3 + group * (group_rows + 2)
        last_detail = first_detail + end_index - first_index - 1
        title = coordinates[first_detail - 1]
        require(text(title.find("s:c", ns)) == "Site" + str(group).zfill(2) and title.get("outlineLevel", "0") == "0", "Group heading/order differs")
        for index in range(first_index, end_index):
            element = coordinates[first_detail + index - first_index]
            cells = element.findall("s:c", ns)
            expected = grouped_values(index)
            require(len(cells) == 7 and element.get("outlineLevel") == "1" and element.get("hidden", "0") == "0", "Detail cell count or outline differs")
            require(text(cells[0]) == "", "The label column overwrote a source value")
            for column, value in enumerate(expected, 1):
                cell = cells[column]
                require(cell.find("s:f", ns) is None, "A source value became a formula")
                if isinstance(value, str):
                    require(cell.get("t") == "inlineStr" and text(cell) == value, "Literal service/site differs at row " + str(index))
                elif value is None:
                    require(cell.find("s:v", ns) is None and cell.find("s:is", ns) is None, "A null metric became a value")
                else:
                    typed = (value - datetime(1899, 12, 30)).total_seconds() / 86400 if isinstance(value, datetime) else value
                    require(abs(number(cell) - typed) < 1e-9, "Typed number/date differs at row " + str(index))
        summary(coordinates[last_detail + 1], first_detail, last_detail, grouped_totals(first_index, end_index), "Subtotal-Site" + str(group).zfill(2))
    summary(rows[-1], 3, last_detail, grouped_totals(0, count), "Grand-total")
    return {"rows": count, "sourceColumns": 6, "columns": 7, "groups": groups, "everyCell": True, "zipCrc": True,
            "typedNumbersNullsUtcDatesAndIds": True, "allNullFirstGroup": True, "numericFormulaCaches": cache_count,
            "finiteSubtotalReferences": True, "nativeOutlines": True, "grandTotals": grouped_totals(0, count), "visibleSheets": 1}


def grouped_pdf(source):
    extracted = subprocess.run([sys.argv[5], "-layout", source, "-"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf8").stdout
    groups = (count + group_rows - 1) // group_rows
    source_lines = [line.strip() for line in extracted.splitlines() if re.match(r"^\s*SRV-\d{6}\s", line)]
    require(len(source_lines) == count, "Grouped PDF lost or duplicated body rows")
    for index, line in enumerate(source_lines):
        cells = line.split()
        expected = grouped_values(index)
        require(len(cells) == 6 and cells[:2] == expected[:2] and cells[5] == "2026-10-20", "PDF row identity/order/date differs")
        for actual, value in zip(cells[2:5], expected[2:5]):
            require(actual == "-" if value is None else abs(float(actual) - value) <= 0.00000051, "PDF metric/null differs at row " + str(index))
    summaries = [line.strip().split() for line in extracted.splitlines() if re.match(r"^\s*(?:Subtotal-Site\d{2}|Grand-total)\s", line)]
    require(len(summaries) == groups + 1, "Grouped PDF subtotal/grand count differs")
    for group, cells in enumerate(summaries):
        grand = group == groups
        expected = grouped_totals(0 if grand else group * group_rows, count if grand else min(count, (group + 1) * group_rows))
        require(len(cells) == 7 and cells[0] == ("Grand-total" if grand else "Subtotal-Site" + str(group).zfill(2)), "PDF summary label/order differs")
        require(cells[1:3] == ["-", "-"] and cells[6] == "-", "PDF summary overwrote an untotaled column")
        for actual, value in zip(cells[3:6], expected):
            require(actual == "-" if value is None else abs(float(actual) - value) <= 0.00000051, "PDF group/grand metric differs")
    continued = 0
    pages = [page for page in extracted.split("\f") if page.strip()]
    for page in pages:
        lines = [line.strip() for line in page.splitlines() if line.strip()]
        for position, line in enumerate(lines):
            if re.match(r"^SRV-\d{6}\s", line):
                group = (int(line.split()[0].split("-")[1]) - 1) // group_rows
                caption = "Site" + str(group).zfill(2)
                require(any(value in (caption, caption + " continued") for value in lines[:position]), "A continued PDF body lost its group context")
        continued += sum(bool(re.match(r"^Site\d{2} continued$", line)) for line in lines)
    require(continued > 0, "The multipage fixture did not qualify continuation captions")
    return {"rows": count, "sourceColumns": 6, "columns": 7, "groups": groups, "everyCell": True, "typedMetricTextAndNulls": True,
            "groupSummaries": groups, "grandTotals": grouped_totals(0, count), "pages": len(pages), "continuationCaptions": continued, "groupContextOnEveryBodyPage": True, "tagged": False}


def service_analysis(package):
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    sheets = workbook.findall("s:sheets/s:sheet", ns)
    require([sheet.get("name") for sheet in sheets] == ["Services", "Overview"], "Expected only the two visible source/report sheets")
    require(all(sheet.get("state", "visible") == "visible" for sheet in sheets), "A source sheet was hidden")
    calculation = workbook.find("s:calcPr", ns)
    require(calculation is not None and calculation.get("calcMode") == "auto" and calculation.get("fullCalcOnLoad") in ("1", "true"), "Native automatic recalculation was not requested")
    services = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
    rows = services.findall("s:sheetData/s:row", ns)
    require(len(rows) == count + 3, "Service title/header/body/footer row count differs")
    require(text(rows[0].find("s:c", ns)) == "Service health report", "Service title differs")
    require([text(cell) for cell in rows[1].findall("s:c", ns)] == ["Service", "Site", "Status", "Latency (ms)", "Availability", "Checked (UTC)"], "Service headings differ")
    statuses = ["Healthy", "Attention", "Offline"]
    totals = dict.fromkeys(statuses, 0)
    for index, element in enumerate(rows[2:-1]):
        require(int(element.get("r")) == index + 3, "Data-row coordinates differ")
        cells = element.findall("s:c", ns)
        require(len(cells) == 6, "Service gained or lost a cell")
        status = "Offline" if index % 37 == 0 else "Attention" if index % 11 == 0 else "Healthy"
        expected = ["WEB-" + str(index + 1).zfill(6), ["Warsaw", "Łódź", "Praha", "Berlin"][index % 4], status]
        require([text(cell) for cell in cells[:3]] == expected and all(cell.get("t") == "inlineStr" for cell in cells[:3]), "Literal identity/site/status differs at row " + str(index))
        totals[status] += 1
        serial = (datetime(2026, 10, 19, 8, index % 60) - datetime(1899, 12, 30)).total_seconds() / 86400
        require(abs(number(cells[3]) - (index % 97 / 10 + 4.75)) < 1e-10 and abs(number(cells[4]) - (0.995 - (index % 17) / 10000)) < 1e-12 and abs(number(cells[5]) - serial) < 1e-9, "Typed latency/availability/UTC date differs at row " + str(index))
    require(int(rows[-1].get("r")) == count + 3 and text(rows[-1].find("s:c", ns)) == "Services", "Footer boundary differs")
    selected = "C3:C" + str(count + 2)
    rules = services.findall("s:dataValidations/s:dataValidation", ns)
    require(len(rules) == 1 and rules[0].get("type") == "list" and rules[0].get("sqref") == selected and rules[0].findtext("s:formula1", "", ns) == '"Healthy,Attention,Offline"' and rules[0].get("showErrorMessage") in ("1", "true"), "Live status validation differs")
    formats = services.findall("s:conditionalFormatting", ns)
    require(len(formats) == 3 and all(item.get("sqref") == selected for item in formats), "Live status formatting range differs")
    require([item.find("s:cfRule", ns).findtext("s:formula", "", ns) for item in formats] == ['"' + status + '"' for status in statuses], "Status formatting literals differ")
    table = ET.fromstring(package.read("xl/tables/table1.xml"))
    require(table.get("name") == "Services" and table.get("ref") == "A2:F" + str(count + 3), "Visible source table range differs")
    overview = ET.fromstring(package.read("xl/worksheets/sheet2.xml"))
    summary = overview.findall("s:sheetData/s:row", ns)
    require(len(summary) == 4, "Summary must retain only its heading and three data rows")
    source_range = "'Services'!$C$3:$C$" + str(count + 2)
    for status, row in zip(statuses, summary[1:]):
        cells = row.findall("s:c", ns)
        require(len(cells) == 2 and text(cells[0]) == status, "Visible summary category differs")
        require(cells[1].findtext("s:f", "", ns) == 'COUNTIF(' + source_range + ',"' + status + '")' and number(cells[1]) == totals[status], "Finite count formula/cache differs")
    chart_files = [name for name in package.namelist() if name.startswith("xl/charts/chart") and name.endswith(".xml")]
    require(len(chart_files) == 1, "Expected exactly one native visible-source chart")
    chart = ET.fromstring(package.read(chart_files[0]))
    require([item.text for item in chart.findall(".//c:f", ns)] == ["'Overview'!$B$1", "'Overview'!$A$2:$A$4", "'Overview'!$B$2:$B$4"], "Chart references are not finite visible summary cells")
    require([item.text for item in chart.findall(".//c:cat/c:strRef/c:strCache/c:pt/c:v", ns)] == statuses, "Chart categories differ")
    points = chart.findall(".//c:numCache/c:pt", ns)
    require(len(points) == 3 and [float(point.findtext("c:v", "", ns)) for point in points] == [totals[status] for status in statuses], "Bound three-point cache differs")
    return {"rows": count, "columns": 6, "everyCell": True, "zipCrc": True, "typedNumbersUtcDatesAndIds": True,
            "summaryCounts": totals, "summaryRows": 3, "formulaCaches": 3, "boundChartPoints": 3, "visibleSheets": 2,
            "finiteSourceRange": source_range, "noHiddenDuplicateSource": True, "liveStatusValidationAndFormatting": True}


if mode == "grouped-pdf":
    print(json.dumps(grouped_pdf(source)))
    sys.exit(0)

with zipfile.ZipFile(source) as package:
    bad = package.testzip()
    if bad:
        raise ValueError("ZIP CRC differs: " + bad)
    if mode == "service-analysis":
        proof = service_analysis(package)
    elif mode == "grouped-analysis":
        proof = grouped_analysis(package)
    elif mode == "inventory":
        document = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
        rows = document.findall("s:sheetData/s:row", ns)
        if len(rows) != count + 1:
            raise ValueError("Worksheet row count differs")
        if ["".join(cell.itertext()) for cell in rows[0].findall("s:c", ns)] != ["ID", "Asset", "Owner", "Status", "Amount", "Seen"]:
            raise ValueError("Worksheet headings differ")
        for index, element in enumerate(rows[1:]):
            expected = ["N" + str(index).zfill(6), "node-" + str(index).zfill(6), "ops-" + str(index % 17),
                        ["Pass", "Warning", "Critical"][index % 3], str(index + 0.5), str(46023 + index % 28)]
            cells = element.findall("s:c", ns)
            actual = ["".join(cell.itertext()) for cell in cells]
            if actual != expected:
                raise ValueError("Worksheet value/order differs at row " + str(index))
            if cells[4].get("t") is not None or cells[5].get("t") is not None or cells[0].get("t") != "inlineStr":
                raise ValueError("Numeric/date or literal identifier cell type differs")
        proof = {"rows": count, "columns": 6, "everyCell": True, "zipCrc": True, "typedAmountDateAndId": True}
    else:
        raise ValueError("Unknown workload validation mode: " + mode)
print(json.dumps(proof))
