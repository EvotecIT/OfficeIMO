"""Independent ZIP/XML readback for the opt-in deterministic workload."""
import json
import sys
import zipfile
import xml.etree.ElementTree as ET

source, count = sys.argv[1], int(sys.argv[2])
ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
with zipfile.ZipFile(source) as package:
    bad = package.testzip()
    if bad:
        raise ValueError("ZIP CRC differs: " + bad)
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
print(json.dumps({"rows": count, "columns": 6, "everyCell": True, "zipCrc": True, "typedAmountDateAndId": True}))
