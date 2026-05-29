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
        private static string ToVisioString(double value) {
            string text = Math.Round(value, 15).ToString("F15", CultureInfo.InvariantCulture);
            return text.TrimEnd('0').TrimEnd('.');
        }

        private static void WriteCell(XmlWriter writer, string ns, string name, double value) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", ToVisioString(value));
            writer.WriteEndElement();
        }

        private static void WriteCell(XmlWriter writer, string ns, string name, double value, string? unit, string? formula) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", ToVisioString(value));
            if (!string.IsNullOrEmpty(unit)) writer.WriteAttributeString("U", unit);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteCellValue(XmlWriter writer, string ns, string name, string value) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", value);
            writer.WriteEndElement();
        }

        private static void WriteCellValue(XmlWriter writer, string ns, string name, string value, string? unit, string? formula) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", value);
            if (!string.IsNullOrEmpty(unit)) writer.WriteAttributeString("U", unit);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteStringCell(XmlWriter writer, string ns, string name, string value, string? formula = null) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", value);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

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

        private static void WriteAttribute(XmlWriter writer, XAttribute attribute) {
            XNamespace attributeNamespace = attribute.Name.Namespace;
            string? namespaceName = attributeNamespace == XNamespace.None ? null : attributeNamespace.NamespaceName;
            writer.WriteAttributeString(null, attribute.Name.LocalName, namespaceName, attribute.Value);
        }

        private static void WriteGeometryHeaderRow(XmlWriter writer, string ns) {
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "Geometry");
            WriteCellValue(writer, ns, "NoFill", "0");
            WriteCellValue(writer, ns, "NoLine", "0");
            WriteCellValue(writer, ns, "NoShow", "0");
            WriteCellValue(writer, ns, "NoSnap", "0");
            WriteCellValue(writer, ns, "NoQuickDrag", "0");
            writer.WriteEndElement();
        }

        private static void WriteXForm(XmlWriter writer, string ns, VisioShape shape, double width, double height) {
            WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, shape.LocPinX, shape.LocPinY, shape.Angle);
        }

        private static void WriteXForm(XmlWriter writer, string ns, double pinX, double pinY, double width, double height, double locPinX, double locPinY, double angle) {
            WriteCell(writer, ns, "PinX", pinX);
            WriteCell(writer, ns, "PinY", pinY);
            WriteCell(writer, ns, "Width", width);
            WriteCell(writer, ns, "Height", height);
            WriteCell(writer, ns, "LocPinX", locPinX);
            WriteCell(writer, ns, "LocPinY", locPinY);
            WriteCell(writer, ns, "Angle", angle);
        }

        private static void WriteXForm1D(XmlWriter writer, string ns, double beginX, double beginY, double endX, double endY) {
            writer.WriteStartElement("XForm1D", ns);
            writer.WriteElementString("BeginX", ns, ToVisioString(beginX));
            writer.WriteElementString("BeginY", ns, ToVisioString(beginY));
            writer.WriteElementString("EndX", ns, ToVisioString(endX));
            writer.WriteElementString("EndY", ns, ToVisioString(endY));
            writer.WriteEndElement();

            WriteCell(writer, ns, "BeginX", beginX);
            WriteCell(writer, ns, "BeginY", beginY);
            WriteCell(writer, ns, "EndX", endX);
            WriteCell(writer, ns, "EndY", endY);
        }

        private static void WriteRectangleGeometry(XmlWriter writer, string ns, double width, double height) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteEllipseGeometry(XmlWriter writer, string ns, double width, double height) {
            double rx = width / 2.0;
            double ry = height / 2.0;
            double centerX = rx;
            double centerY = ry;
            const int segments = 24;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", centerX + rx);
            WriteCell(writer, ns, "Y", centerY);
            writer.WriteEndElement();

            for (int i = 1; i <= segments; i++) {
                double angle = (Math.PI * 2D * i) / segments;
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("T", "LineTo");
                WriteCell(writer, ns, "X", centerX + (Math.Cos(angle) * rx));
                WriteCell(writer, ns, "Y", centerY + (Math.Sin(angle) * ry));
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteDiamondGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double midY = height / 2.0;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteTriangleGeometry(XmlWriter writer, string ns, double width, double height) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width / 2.0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WritePentagonGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double shoulderY = height * 0.62;
            double lowerInset = width * 0.2;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - lowerInset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", lowerInset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteParallelogramGeometry(XmlWriter writer, string ns, double width, double height) {
            double offset = Math.Min(width / 4.0, Math.Max(width / 10.0, height / 3.0));
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", offset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - offset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", offset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteHexagonGeometry(XmlWriter writer, string ns, double width, double height) {
            double inset = Math.Min(width / 4.0, Math.Max(width / 8.0, height / 4.0));
            double midY = height / 2.0;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - inset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteTrapezoidGeometry(XmlWriter writer, string ns, double width, double height) {
            double inset = Math.Min(width / 5.0, Math.Max(width / 10.0, height / 4.0));
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteOffPageReferenceGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double shoulderY = height * 0.45;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

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

        private static void WritePageCell(XmlWriter writer, string ns, string name, double value, string? unit = null, string? formula = null) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", XmlConvert.ToString(value));
            if (unit != null) writer.WriteAttributeString("U", unit);
            if (formula != null) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteTextElement(XmlWriter writer, string ns, string? text, XElement? preservedTextElement = null, string? preservedTextValue = null) {
            if (preservedTextElement != null &&
                string.Equals(text ?? string.Empty, preservedTextValue ?? string.Empty, StringComparison.Ordinal)) {
                XElement clone = new(preservedTextElement);
                using var reader = clone.CreateReader();
                writer.WriteNode(reader, false);
                return;
            }

            if (!string.IsNullOrEmpty(text)) {
                writer.WriteElementString("Text", ns, text);
            }
        }

        private static string GetConnectionCell(VisioShape shape, VisioConnectionPoint? point, string? preservedCell = null) {
            if (point == null) {
                return string.IsNullOrWhiteSpace(preservedCell) ? "PinX" : preservedCell!;
            }

            Dictionary<VisioConnectionPoint, int> pointIndices = BuildConnectionPointIndices(shape.ConnectionPoints);
            return pointIndices.TryGetValue(point, out int index)
                ? $"Connections.X{index + 1}"
                : string.IsNullOrWhiteSpace(preservedCell) ? "PinX" : preservedCell!;
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

        private static XDocument CreateVisioDocumentXml(
            bool requestRecalcOnOpen,
            IEnumerable<XAttribute>? preservedDocumentAttributes = null,
            IEnumerable<XElement>? preservedDocumentElements = null,
            IEnumerable<XAttribute>? preservedDocumentSettingsAttributes = null,
            IEnumerable<XElement>? preservedDocumentSettingsElements = null,
            IEnumerable<XAttribute>? preservedColorsAttributes = null,
            IEnumerable<XElement>? preservedColorsElements = null,
            IEnumerable<XAttribute>? preservedFaceNamesAttributes = null,
            IEnumerable<XElement>? preservedFaceNamesElements = null,
            IEnumerable<XAttribute>? preservedStyleSheetsAttributes = null,
            IEnumerable<XElement>? preservedStyleSheetsElements = null,
            IDictionary<string, PreservedStyleSheetData>? preservedGeneratedStyleSheets = null,
            IEnumerable<XElement>? preservedAdditionalStyleSheets = null) {
            XNamespace ns = VisioNamespace;
            XElement settings = new(ns + "DocumentSettings",
                new XAttribute("TopPage", 0),
                new XAttribute("DefaultTextStyle", 0),
                new XAttribute("DefaultLineStyle", 0),
                new XAttribute("DefaultFillStyle", 0),
                new XAttribute("DefaultGuideStyle", 4),
                new XElement(ns + "GlueSettings", 9),
                new XElement(ns + "SnapSettings", 295),
                new XElement(ns + "SnapExtensions", 34),
                new XElement(ns + "SnapAngles"),
                new XElement(ns + "DynamicGridEnabled", 1),
                new XElement(ns + "ProtectStyles", 0),
                new XElement(ns + "ProtectShapes", 0),
                new XElement(ns + "ProtectMasters", 0),
                new XElement(ns + "ProtectBkgnds", 0));
            foreach (XAttribute attribute in preservedDocumentSettingsAttributes ?? Enumerable.Empty<XAttribute>()) {
                settings.Add(new XAttribute(attribute));
            }
            if (requestRecalcOnOpen) settings.Add(new XElement(ns + "RelayoutAndRerouteUponOpen", 1));
            foreach (XElement element in preservedDocumentSettingsElements ?? Enumerable.Empty<XElement>()) {
                settings.Add(new XElement(element));
            }
            XElement colors = new(ns + "Colors");
            foreach (XAttribute attribute in preservedColorsAttributes ?? Enumerable.Empty<XAttribute>()) {
                colors.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedColorsElements ?? Enumerable.Empty<XElement>()) {
                colors.Add(new XElement(element));
            }
            XElement faceNames = new(ns + "FaceNames");
            foreach (XAttribute attribute in preservedFaceNamesAttributes ?? Enumerable.Empty<XAttribute>()) {
                faceNames.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedFaceNamesElements ?? Enumerable.Empty<XElement>()) {
                faceNames.Add(new XElement(element));
            }
            XElement styleSheets = new(ns + "StyleSheets");
            foreach (XAttribute attribute in preservedStyleSheetsAttributes ?? Enumerable.Empty<XAttribute>()) {
                styleSheets.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedStyleSheetsElements ?? Enumerable.Empty<XElement>()) {
                styleSheets.Add(new XElement(element));
            }
            styleSheets.Add(CreateGeneratedStyleSheet(ns, "0", preservedGeneratedStyleSheets));
            styleSheets.Add(CreateGeneratedStyleSheet(ns, "1", preservedGeneratedStyleSheets));
            styleSheets.Add(CreateGeneratedStyleSheet(ns, "2", preservedGeneratedStyleSheets));
            foreach (XElement styleSheet in preservedAdditionalStyleSheets ?? Enumerable.Empty<XElement>()) {
                styleSheets.Add(new XElement(styleSheet));
            }

            XElement root = new(ns + "VisioDocument");
            foreach (XAttribute attribute in preservedDocumentAttributes ?? Enumerable.Empty<XAttribute>()) {
                root.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedDocumentElements ?? Enumerable.Empty<XElement>()) {
                root.Add(new XElement(element));
            }
            root.Add(settings);
            root.Add(colors);
            root.Add(faceNames);
            root.Add(styleSheets);

            return new XDocument(root);
        }

        private static XElement CreateGeneratedStyleSheet(XNamespace ns, string styleSheetId, IDictionary<string, PreservedStyleSheetData>? preservedGeneratedStyleSheets) {
            XElement styleSheet = styleSheetId switch {
                "0" => new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 0),
                    new XAttribute("Name", "No Style"),
                    new XAttribute("NameU", "No Style"),
                    new XElement(ns + "Cell", new XAttribute("N", "EnableLineProps"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "EnableFillProps"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "EnableTextProps"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "LineWeight"), new XAttribute("V", "0.01041666666666667")),
                    new XElement(ns + "Cell", new XAttribute("N", "LineColor"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "LinePattern"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "FillForegnd"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "FillPattern"), new XAttribute("V", "1"))),
                "1" => new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 1),
                    new XAttribute("Name", "Normal"),
                    new XAttribute("NameU", "Normal"),
                    new XAttribute("BasedOn", 0),
                    new XAttribute("LineStyle", 0),
                    new XAttribute("FillStyle", 0),
                    new XAttribute("TextStyle", 0),
                    new XElement(ns + "Cell", new XAttribute("N", "LinePattern"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "LineColor"), new XAttribute("V", "#000000")),
                    new XElement(ns + "Cell", new XAttribute("N", "FillPattern"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "FillForegnd"), new XAttribute("V", "#FFFFFF"))),
                "2" => new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 2),
                    new XAttribute("Name", "Connector"),
                    new XAttribute("NameU", "Connector"),
                    new XAttribute("BasedOn", 1),
                    new XAttribute("LineStyle", 0),
                    new XAttribute("FillStyle", 0),
                    new XAttribute("TextStyle", 0),
                    new XElement(ns + "Cell", new XAttribute("N", "EndArrow"), new XAttribute("V", 0))),
                _ => throw new InvalidOperationException($"Unsupported generated style sheet id '{styleSheetId}'.")
            };

            if (preservedGeneratedStyleSheets != null &&
                preservedGeneratedStyleSheets.TryGetValue(styleSheetId, out PreservedStyleSheetData? preserved)) {
                foreach (XAttribute attribute in preserved.Attributes) {
                    styleSheet.Add(new XAttribute(attribute));
                }

                foreach (XElement element in preserved.ChildElements) {
                    styleSheet.Add(new XElement(element));
                }
            }

            return styleSheet;
        }

        private static void FixContentTypes(string filePath, int masterCount, bool includeTheme, IEnumerable<string> pagePartNames) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be null or whitespace.", nameof(filePath));
            }

            if (pagePartNames is null) {
                throw new ArgumentNullException(nameof(pagePartNames));
            }

            using FileStream zipStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Update);
            FixContentTypesCore(archive, masterCount, includeTheme, pagePartNames);
        }

        private static void FixContentTypes(Stream stream, int masterCount, bool includeTheme, IEnumerable<string> pagePartNames) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }
            if (!stream.CanRead || !stream.CanWrite || !stream.CanSeek) {
                throw new ArgumentException("Stream must be readable, writable, and seekable.", nameof(stream));
            }
            if (pagePartNames is null) {
                throw new ArgumentNullException(nameof(pagePartNames));
            }

            stream.Seek(0, SeekOrigin.Begin);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update, leaveOpen: true);
            FixContentTypesCore(archive, masterCount, includeTheme, pagePartNames);
            stream.Seek(0, SeekOrigin.Begin);
        }

        private static void FixContentTypesCore(ZipArchive archive, int masterCount, bool includeTheme, IEnumerable<string> pagePartNames) {
            ZipArchiveEntry? entry = archive.GetEntry("[Content_Types].xml");
            entry?.Delete();
            ZipArchiveEntry newEntry = archive.CreateEntry("[Content_Types].xml");
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement root = new(ct + "Types",
                new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "emf"), new XAttribute("ContentType", "image/x-emf")),
                new XElement(ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")),
                new XElement(ct + "Default", new XAttribute("Extension", "jpg"), new XAttribute("ContentType", "image/jpeg")),
                new XElement(ct + "Default", new XAttribute("Extension", "jpeg"), new XAttribute("ContentType", "image/jpeg")),
                new XElement(ct + "Default", new XAttribute("Extension", "gif"), new XAttribute("ContentType", "image/gif")),
                new XElement(ct + "Default", new XAttribute("Extension", "svg"), new XAttribute("ContentType", "image/svg+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "tif"), new XAttribute("ContentType", "image/tiff")),
                new XElement(ct + "Default", new XAttribute("Extension", "tiff"), new XAttribute("ContentType", "image/tiff")));

            HashSet<string> overridePartNames = new(StringComparer.OrdinalIgnoreCase);
            void AddOverride(string partName, string contentType) {
                if (string.IsNullOrWhiteSpace(partName)) {
                    return;
                }

                string normalizedPartName = NormalizePartName(partName);

                if (overridePartNames.Add(normalizedPartName)) {
                    root.Add(new XElement(ct + "Override",
                        new XAttribute("PartName", normalizedPartName),
                        new XAttribute("ContentType", contentType)));
                }
            }

            AddOverride("/visio/document.xml", DocumentContentType);
            AddOverride("/visio/pages/pages.xml", PagesContentType);
            AddOverride("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
            AddOverride("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            AddOverride("/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
            AddOverride("/docProps/thumbnail.emf", "image/x-emf");
            AddOverride("/visio/windows.xml", WindowsContentType);

            foreach (string partName in pagePartNames) {
                AddOverride(partName, PageContentType);
            }
            if (includeTheme) {
                AddOverride("/visio/theme/theme1.xml", ThemeContentType);
            }
            if (masterCount > 0) {
                AddOverride("/visio/masters/masters.xml", "application/vnd.ms-visio.masters+xml");
                for (int i = 1; i <= masterCount; i++) {
                    AddOverride($"/visio/masters/master{i}.xml", "application/vnd.ms-visio.master+xml");
                }
            }
            XDocument doc = new(root);
            using StreamWriter writer = new(newEntry.Open());
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static string NormalizePartName(string partName) {
            if (partName is null) {
                throw new ArgumentNullException(nameof(partName));
            }

            return "/" + partName.TrimStart('/');
        }
    }
}
