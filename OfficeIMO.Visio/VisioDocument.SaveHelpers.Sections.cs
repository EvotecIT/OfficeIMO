using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Save-time helper methods for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {

        private static void WriteHyperlinkSection(XmlWriter writer, string ns, IList<VisioHyperlink> hyperlinks) {
            if (hyperlinks.Count == 0) {
                return;
            }

            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Hyperlink");
            for (int i = 0; i < hyperlinks.Count; i++) {
                VisioHyperlink hyperlink = hyperlinks[i];
                writer.WriteStartElement("Row", ns);
                if (!string.IsNullOrWhiteSpace(hyperlink.RowName)) {
                    writer.WriteAttributeString("N", hyperlink.RowName);
                } else if (hyperlink.RowIndex.HasValue) {
                    writer.WriteAttributeString("IX", hyperlink.RowIndex.Value.ToString(CultureInfo.InvariantCulture));
                } else {
                    writer.WriteAttributeString("N", "Row_" + (i + 1).ToString(CultureInfo.InvariantCulture));
                }

                foreach (XAttribute attribute in hyperlink.PreservedRowAttributes) {
                    WriteAttribute(writer, attribute);
                }

                foreach (string cellName in VisioHyperlink.CellOrder) {
                    WriteHyperlinkCell(writer, ns, hyperlink, cellName);
                }

                foreach (XElement cell in hyperlink.PreservedCells) {
                    cell.WriteTo(writer);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteUserSection(XmlWriter writer, string ns, IList<VisioUserCell> userCells) {
            if (userCells.Count == 0) {
                return;
            }

            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "User");
            for (int i = 0; i < userCells.Count; i++) {
                VisioUserCell userCell = userCells[i];
                writer.WriteStartElement("Row", ns);
                if (!string.IsNullOrWhiteSpace(userCell.Name)) {
                    writer.WriteAttributeString("N", userCell.Name);
                } else if (userCell.RowIndex.HasValue) {
                    writer.WriteAttributeString("IX", userCell.RowIndex.Value.ToString(CultureInfo.InvariantCulture));
                } else {
                    writer.WriteAttributeString("N", "Row_" + (i + 1).ToString(CultureInfo.InvariantCulture));
                }

                foreach (XAttribute attribute in userCell.PreservedRowAttributes) {
                    WriteAttribute(writer, attribute);
                }

                writer.WriteStartElement("Cell", ns);
                writer.WriteAttributeString("N", "Value");
                writer.WriteAttributeString("V", userCell.Value ?? string.Empty);
                if (!string.IsNullOrEmpty(userCell.Unit)) writer.WriteAttributeString("U", userCell.Unit);
                if (!string.IsNullOrEmpty(userCell.Formula)) writer.WriteAttributeString("F", userCell.Formula);
                foreach (XAttribute attribute in userCell.PreservedValueAttributes) {
                    WriteAttribute(writer, attribute);
                }
                writer.WriteEndElement();

                if (userCell.Prompt != null ||
                    !string.IsNullOrEmpty(userCell.PromptFormula) ||
                    userCell.PreservedPromptAttributes.Count > 0) {
                    writer.WriteStartElement("Cell", ns);
                    writer.WriteAttributeString("N", "Prompt");
                    writer.WriteAttributeString("V", userCell.Prompt ?? string.Empty);
                    if (!string.IsNullOrEmpty(userCell.PromptFormula)) writer.WriteAttributeString("F", userCell.PromptFormula);
                    foreach (XAttribute attribute in userCell.PreservedPromptAttributes) {
                        WriteAttribute(writer, attribute);
                    }
                    writer.WriteEndElement();
                }

                foreach (XElement cell in userCell.PreservedCells) {
                    cell.WriteTo(writer);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteHyperlinkCell(XmlWriter writer, string ns, VisioHyperlink hyperlink, string cellName) {
            string value = cellName switch {
                "Description" => hyperlink.Description ?? string.Empty,
                "Address" => hyperlink.Address ?? string.Empty,
                "SubAddress" => hyperlink.SubAddress ?? string.Empty,
                "ExtraInfo" => hyperlink.ExtraInfo ?? string.Empty,
                "Frame" => hyperlink.Frame ?? string.Empty,
                "NewWindow" => hyperlink.NewWindow ? "1" : "0",
                "Default" => hyperlink.Default ? "1" : "0",
                "Invisible" => hyperlink.Invisible ? "1" : "0",
                "SortKey" => hyperlink.SortKey ?? string.Empty,
                _ => string.Empty
            };

            hyperlink.PreservedKnownCells.TryGetValue(cellName, out XElement? template);
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", cellName);
            writer.WriteAttributeString("V", value);
            if (template != null) {
                foreach (XAttribute attribute in template.Attributes()) {
                    if (attribute.IsNamespaceDeclaration ||
                        string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "V", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    WriteAttribute(writer, attribute);
                }
            }

            writer.WriteEndElement();
        }

        private static void WriteConnectionSection(XmlWriter writer, string ns, IList<VisioConnectionPoint> points) {
            if (points.Count == 0) return;
            Dictionary<VisioConnectionPoint, int> pointIndices = BuildConnectionPointIndices(points);
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Connection");
            for (int i = 0; i < points.Count; i++) {
                VisioConnectionPoint cp = points[i];
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("T", "Connection");
                writer.WriteAttributeString("IX", XmlConvert.ToString(pointIndices[cp]));
                WriteCell(writer, ns, "X", cp.X);
                WriteCell(writer, ns, "Y", cp.Y);
                WriteCell(writer, ns, "DirX", cp.DirX);
                WriteCell(writer, ns, "DirY", cp.DirY);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WriteDataSection(XmlWriter writer, string ns, IDictionary<string, string> data, IEnumerable<XElement>? preservedRows = null, KeyValuePair<string, string>? additionalEntry = null, IList<VisioShapeDataRow>? shapeDataRows = null) {
            if (data.Count == 0 && additionalEntry == null && (shapeDataRows == null || shapeDataRows.Count == 0)) return;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Prop");

            HashSet<string> emittedKeys = new(StringComparer.Ordinal);
            if (shapeDataRows != null) {
                foreach (VisioShapeDataRow row in shapeDataRows) {
                    WriteShapeDataRow(writer, ns, row, data);
                    emittedKeys.Add(row.Name);
                }
            }

            if (preservedRows != null) {
                foreach (XElement preservedRow in preservedRows) {
                    if (!(preservedRow.Attribute("N")?.Value is string key) ||
                        key.Length == 0 ||
                        string.Equals(key, OriginalIdPropName, StringComparison.Ordinal) ||
                        emittedKeys.Contains(key)) {
                        continue;
                    }

                    if (!data.TryGetValue(key, out string? value)) {
                        continue;
                    }

                    XElement clone = new(preservedRow);
                    XElement? valueCell = clone.Elements(XName.Get("Cell", clone.Name.NamespaceName))
                        .FirstOrDefault(cell => string.Equals(cell.Attribute("N")?.Value, "Value", StringComparison.Ordinal));
                    if (valueCell == null) {
                        valueCell = new XElement(XName.Get("Cell", clone.Name.NamespaceName),
                            new XAttribute("N", "Value"));
                        clone.Add(valueCell);
                    }
                    valueCell.SetAttributeValue("V", value);
                    valueCell.Attribute("F")?.Remove();

                    using var reader = clone.CreateReader();
                    writer.WriteNode(reader, false);
                    emittedKeys.Add(key);
                }
            }

            foreach (var kv in data) {
                if (emittedKeys.Contains(kv.Key)) {
                    continue;
                }

                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("N", kv.Key);
                writer.WriteStartElement("Cell", ns);
                writer.WriteAttributeString("N", "Value");
                writer.WriteAttributeString("V", kv.Value);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            if (additionalEntry.HasValue) {
                KeyValuePair<string, string> extra = additionalEntry.Value;
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("N", extra.Key);
                writer.WriteStartElement("Cell", ns);
                writer.WriteAttributeString("N", "Value");
                writer.WriteAttributeString("V", extra.Value);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WriteShapeDataRow(XmlWriter writer, string ns, VisioShapeDataRow row, IDictionary<string, string> data) {
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("N", row.Name);
            if (row.RowIndex.HasValue) {
                writer.WriteAttributeString("IX", row.RowIndex.Value.ToString(CultureInfo.InvariantCulture));
            }

            foreach (XAttribute attribute in row.PreservedRowAttributes) {
                WriteAttribute(writer, attribute);
            }

            foreach (string cellName in GetShapeDataCellWriteOrder(row)) {
                WriteShapeDataCell(writer, ns, row, data, cellName);
            }

            foreach (XElement cell in row.PreservedCells) {
                cell.WriteTo(writer);
            }

            writer.WriteEndElement();
        }

        private static IEnumerable<string> GetShapeDataCellWriteOrder(VisioShapeDataRow row) {
            HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);
            foreach (string cellName in row.PreservedCellOrder) {
                if (emitted.Add(cellName)) {
                    yield return cellName;
                }
            }

            foreach (string cellName in VisioShapeDataRow.CellOrder) {
                if (emitted.Add(cellName)) {
                    yield return cellName;
                }
            }
        }

        private static void WriteShapeDataCell(XmlWriter writer, string ns, VisioShapeDataRow row, IDictionary<string, string> data, string cellName) {
            row.PreservedKnownCells.TryGetValue(cellName, out XElement? template);
            string? value = GetShapeDataCellValue(row, data, cellName, template);
            string? formula = GetShapeDataCellFormula(row, data, cellName, template);
            string? unit = cellName == "Value" ? row.ValueUnit ?? template?.Attribute("U")?.Value : template?.Attribute("U")?.Value;
            if (value == null && formula == null && template == null) {
                return;
            }

            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", cellName);
            if (value != null) writer.WriteAttributeString("V", value);
            if (!string.IsNullOrEmpty(unit)) writer.WriteAttributeString("U", unit);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            if (template != null) {
                foreach (XAttribute attribute in template.Attributes()) {
                    if (attribute.IsNamespaceDeclaration ||
                        string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "V", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "U", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "F", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    WriteAttribute(writer, attribute);
                }
            }

            writer.WriteEndElement();
        }

        private static string? GetShapeDataCellValue(VisioShapeDataRow row, IDictionary<string, string> data, string cellName, XElement? template) {
            switch (cellName) {
                case "Value":
                    if (IsDictionaryShapeDataOverride(row, data, out string? currentValue)) {
                        return currentValue;
                    }

                    return row.Value ?? template?.Attribute("V")?.Value;
                case "Label":
                    return row.Label ?? template?.Attribute("V")?.Value;
                case "Prompt":
                    return row.Prompt ?? template?.Attribute("V")?.Value;
                case "Type":
                    return row.Type.HasValue ? ((int)row.Type.Value).ToString(CultureInfo.InvariantCulture) : template?.Attribute("V")?.Value;
                case "Format":
                    return row.Format ?? template?.Attribute("V")?.Value;
                case "SortKey":
                    return row.SortKey ?? template?.Attribute("V")?.Value;
                case "Invisible":
                    return row.Invisible.HasValue ? (row.Invisible.Value ? "1" : "0") : template?.Attribute("V")?.Value;
                case "Verify":
                    return row.Verify.HasValue ? (row.Verify.Value ? "1" : "0") : template?.Attribute("V")?.Value;
                case "DataLinked":
                    return row.DataLinked.HasValue ? (row.DataLinked.Value ? "1" : "0") : template?.Attribute("V")?.Value;
                case "Calendar":
                    return row.Calendar ?? template?.Attribute("V")?.Value;
                case "LangID":
                    return row.LangId ?? template?.Attribute("V")?.Value;
                default:
                    return template?.Attribute("V")?.Value;
            }
        }

        private static string? GetShapeDataCellFormula(VisioShapeDataRow row, IDictionary<string, string> data, string cellName, XElement? template) {
            if (string.Equals(cellName, "Value", StringComparison.OrdinalIgnoreCase) &&
                (IsDictionaryShapeDataOverride(row, data, out _) ||
                 (row.LoadedValue != null &&
                  !string.Equals(row.Value, row.LoadedValue, StringComparison.Ordinal) &&
                  string.Equals(row.ValueFormula, template?.Attribute("F")?.Value, StringComparison.Ordinal)))) {
                return null;
            }

            string? formula = cellName switch {
                "Value" => row.ValueFormula,
                "Label" => row.LabelFormula,
                "Prompt" => row.PromptFormula,
                "Type" => row.TypeFormula,
                "Format" => row.FormatFormula,
                "SortKey" => row.SortKeyFormula,
                "Invisible" => row.InvisibleFormula,
                "Verify" => row.VerifyFormula,
                "DataLinked" => row.DataLinkedFormula,
                "Calendar" => row.CalendarFormula,
                "LangID" => row.LangIdFormula,
                _ => null
            };

            return formula ?? template?.Attribute("F")?.Value;
        }

        private static bool IsDictionaryShapeDataOverride(VisioShapeDataRow row, IDictionary<string, string> data, out string? value) {
            value = null;
            if (data.TryGetValue(row.Name, out string? current) &&
                !string.Equals(current, row.MirroredDataValue, StringComparison.Ordinal) &&
                !string.Equals(current, row.Value, StringComparison.Ordinal)) {
                value = current;
                return true;
            }

            return false;
        }

        private static Dictionary<VisioConnectionPoint, int> BuildConnectionPointIndices(IList<VisioConnectionPoint> points) {
            Dictionary<VisioConnectionPoint, int> indices = new(points.Count);
            HashSet<int> usedIndices = new();

            foreach (VisioConnectionPoint point in points) {
                if (point.SectionIndex.HasValue && point.SectionIndex.Value >= 0 && usedIndices.Add(point.SectionIndex.Value)) {
                    indices[point] = point.SectionIndex.Value;
                }
            }

            int nextIndex = 0;
            foreach (VisioConnectionPoint point in points) {
                if (indices.ContainsKey(point)) {
                    continue;
                }

                while (usedIndices.Contains(nextIndex)) {
                    nextIndex++;
                }

                indices[point] = nextIndex;
                usedIndices.Add(nextIndex);
                nextIndex++;
            }

            return indices;
        }
    }
}
