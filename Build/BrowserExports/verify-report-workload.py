"""Independent ZIP/XML readback for the opt-in deterministic workload."""
import json
import sys
import zipfile
import math
from datetime import datetime
import xml.etree.ElementTree as ET

source, count = sys.argv[1], int(sys.argv[2])
mode = sys.argv[3] if len(sys.argv) > 3 else "inventory"
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


with zipfile.ZipFile(source) as package:
    bad = package.testzip()
    if bad:
        raise ValueError("ZIP CRC differs: " + bad)
    if mode == "service-analysis":
        proof = service_analysis(package)
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
