using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private static void ParseShapeDataRows(XElement propSection, XNamespace ns, VisioShape shape) {
            ParseShapeDataRows(propSection, ns, shape.ShapeData, shape.PreservedDataRows, shape.Data);
        }

        private static void ParseShapeDataRows(
            XElement propSection,
            XNamespace ns,
            IList<VisioShapeDataRow> shapeData,
            IList<XElement> preservedDataRows,
            IDictionary<string, string> data) {
            shapeData.Clear();
            preservedDataRows.Clear();
            foreach (XElement row in propSection.Elements(ns + "Row")) {
                string? key = row.Attribute("N")?.Value;
                if (string.IsNullOrEmpty(key) || string.Equals(key, OriginalIdPropName, StringComparison.Ordinal)) {
                    continue;
                }

                VisioShapeDataRow dataRow = new(key!);
                if (int.TryParse(row.Attribute("IX")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rowIndex) &&
                    rowIndex >= 0) {
                    dataRow.RowIndex = rowIndex;
                }

                foreach (XAttribute attribute in row.Attributes()) {
                    if (attribute.IsNamespaceDeclaration ||
                        string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "IX", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    dataRow.PreservedRowAttributes.Add(new XAttribute(attribute));
                }

                foreach (XElement cell in row.Elements(ns + "Cell")) {
                    string? cellName = cell.Attribute("N")?.Value;
                    string? value = cell.Attribute("V")?.Value;
                    if (IsShapeDataCell(cellName)) {
                        dataRow.PreservedKnownCells[cellName!] = new XElement(cell);
                        dataRow.PreservedCellOrder.Add(cellName!);
                        ApplyShapeDataCell(dataRow, cellName!, value, cell.Attribute("F")?.Value, cell.Attribute("U")?.Value);
                    } else {
                        dataRow.PreservedCells.Add(new XElement(cell));
                    }
                }

                dataRow.LoadedValue = dataRow.Value;
                if (dataRow.Value != null) {
                    data[dataRow.Name] = dataRow.Value;
                    dataRow.MirroredDataValue = dataRow.Value;
                }

                shapeData.Add(dataRow);
            }
        }

        private static bool IsShapeDataCell(string? cellName) {
            if (string.IsNullOrEmpty(cellName)) {
                return false;
            }

            return VisioShapeDataRow.CellOrder.Any(current => string.Equals(current, cellName, StringComparison.OrdinalIgnoreCase));
        }

        private static void ApplyShapeDataCell(VisioShapeDataRow row, string cellName, string? value, string? formula, string? unit) {
            switch (cellName) {
                case "Value":
                    row.Value = value;
                    row.ValueFormula = formula;
                    row.ValueUnit = unit;
                    break;
                case "Label":
                    row.Label = value;
                    row.LabelFormula = formula;
                    break;
                case "Prompt":
                    row.Prompt = value;
                    row.PromptFormula = formula;
                    break;
                case "Type":
                    if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int typeValue) &&
                        Enum.IsDefined(typeof(VisioShapeDataType), typeValue)) {
                        row.Type = (VisioShapeDataType)typeValue;
                    }
                    row.TypeFormula = formula;
                    break;
                case "Format":
                    row.Format = value;
                    row.FormatFormula = formula;
                    break;
                case "SortKey":
                    row.SortKey = value;
                    row.SortKeyFormula = formula;
                    break;
                case "Invisible":
                    row.Invisible = ParseNullableBoolCell(value);
                    row.InvisibleFormula = formula;
                    break;
                case "Verify":
                    row.Verify = ParseNullableBoolCell(value);
                    row.VerifyFormula = formula;
                    break;
                case "DataLinked":
                    row.DataLinked = ParseNullableBoolCell(value);
                    row.DataLinkedFormula = formula;
                    break;
                case "Calendar":
                    row.Calendar = value;
                    row.CalendarFormula = formula;
                    break;
                case "LangID":
                    row.LangId = value;
                    row.LangIdFormula = formula;
                    break;
            }
        }

        private static bool? ParseNullableBoolCell(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intValue)) {
                return intValue != 0;
            }

            if (bool.TryParse(value, out bool boolValue)) {
                return boolValue;
            }

            return null;
        }
    }
}
