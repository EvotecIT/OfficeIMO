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

        private static bool IsConnectorShape(XElement shapeElement, IReadOnlyDictionary<string, VisioMaster> masters) {
            string? nameU = shapeElement.Attribute("NameU")?.Value;
            if (string.Equals(nameU, "Connector", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(nameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (TryGetTruthyCellValue(shapeElement, "OneD")) {
                return true;
            }

            string? masterId = shapeElement.Attribute("Master")?.Value;
            if (!string.IsNullOrEmpty(masterId) &&
                masters.TryGetValue(masterId!, out VisioMaster? master) &&
                string.Equals(master.NameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
            
            return false;
        }

        private static ConnectorKind DetermineConnectorKind(XElement connectorElement, XNamespace ns, IReadOnlyDictionary<string, VisioMaster> masters) {
            if (HasDynamicConnectorIdentity(connectorElement, masters)) {
                return ConnectorKind.Dynamic;
            }

            XElement? geometrySection = connectorElement.Elements(ns + "Section")
                .FirstOrDefault(e => e.Attribute("N")?.Value == "Geometry");
            if (geometrySection == null) {
                return ConnectorKind.Dynamic;
            }

            List<XElement> rows = geometrySection.Elements(ns + "Row").ToList();
            List<XElement> drawableRows = rows
                .Where(row => !string.Equals(row.Attribute("T")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (drawableRows.Count == 0) {
                return ConnectorKind.Dynamic;
            }

            if (drawableRows.Any(IsCurvedGeometryRow)) {
                return ConnectorKind.Curved;
            }

            List<(double X, double Y)> points = new();
            foreach (XElement row in drawableRows) {
                string? type = row.Attribute("T")?.Value;
                if (!string.Equals(type, "MoveTo", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(type, "LineTo", StringComparison.OrdinalIgnoreCase)) {
                    return ConnectorKind.Curved;
                }

                points.Add((GetCellValue(row, ns, "X"), GetCellValue(row, ns, "Y")));
            }

            if (points.Count <= 2) {
                return ConnectorKind.Straight;
            }

            bool allOrthogonal = true;
            for (int i = 1; i < points.Count; i++) {
                (double previousX, double previousY) = points[i - 1];
                (double currentX, double currentY) = points[i];
                bool sameX = Math.Abs(previousX - currentX) <= 1e-9;
                bool sameY = Math.Abs(previousY - currentY) <= 1e-9;
                if (!sameX && !sameY) {
                    allOrthogonal = false;
                    break;
                }
            }

            return allOrthogonal ? ConnectorKind.RightAngle : ConnectorKind.Curved;
        }

        private static void TryHydrateConnectorWaypoints(VisioConnector connector, XElement connectorElement, XNamespace ns) {
            if (connector.Waypoints.Count > 0) {
                return;
            }

            XElement? geometrySection = connectorElement.Elements(ns + "Section")
                .FirstOrDefault(section => string.Equals(section.Attribute("N")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase));
            if (geometrySection == null) {
                return;
            }

            List<XElement> rows = geometrySection.Elements(ns + "Row")
                .Where(row => !string.Equals(row.Attribute("T")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (rows.Count < 3 ||
                !string.Equals(rows[0].Attribute("T")?.Value, "MoveTo", StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            List<(double X, double Y)> points = new();
            for (int i = 0; i < rows.Count; i++) {
                XElement row = rows[i];
                string? rowType = row.Attribute("T")?.Value;
                if (i > 0 && !string.Equals(rowType, "LineTo", StringComparison.OrdinalIgnoreCase)) {
                    return;
                }

                if (!TryGetNumericCellValue(row, ns, "X", out double x) ||
                    !TryGetNumericCellValue(row, ns, "Y", out double y)) {
                    return;
                }

                points.Add((x, y));
            }

            ComputeConnectorEndpoints(connector, out double startX, out double startY, out double endX, out double endY);
            if (!PointsEqual(points[0].X, points[0].Y, startX, startY) ||
                !PointsEqual(points[points.Count - 1].X, points[points.Count - 1].Y, endX, endY)) {
                return;
            }

            for (int i = 1; i < points.Count - 1; i++) {
                connector.Waypoints.Add(new VisioConnectorWaypoint(points[i].X, points[i].Y));
            }
        }

        private static bool TryGetNumericCellValue(XElement row, XNamespace ns, string cellName, out double result) {
            string? value = row.Elements(ns + "Cell")
                .FirstOrDefault(cell => string.Equals(cell.Attribute("N")?.Value, cellName, StringComparison.OrdinalIgnoreCase))
                ?.Attribute("V")?.Value;
            string? normalized = NormalizeCellLiteral(value);
            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out result) &&
                !double.IsNaN(result) &&
                !double.IsInfinity(result)) {
                return true;
            }

            result = 0;
            return false;
        }

        private static bool PointsEqual(double actualX, double actualY, double expectedX, double expectedY) {
            const double tolerance = 1e-7;
            return Math.Abs(actualX - expectedX) <= tolerance &&
                   Math.Abs(actualY - expectedY) <= tolerance;
        }

        private static bool HasDynamicConnectorIdentity(XElement connectorElement, IReadOnlyDictionary<string, VisioMaster> masters) {
            string? nameU = connectorElement.Attribute("NameU")?.Value;
            if (string.Equals(nameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string? masterId = connectorElement.Attribute("Master")?.Value;
            return !string.IsNullOrEmpty(masterId) &&
                   masters.TryGetValue(masterId!, out VisioMaster? master) &&
                   string.Equals(master.NameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCurvedGeometryRow(XElement row) {
            string? type = row.Attribute("T")?.Value;
            if (string.IsNullOrEmpty(type)) {
                return false;
            }

            string rowType = type!;
            return rowType.IndexOf("Arc", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("Spline", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("Bezier", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("NURBS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("Curve", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static double GetCellValue(XElement row, XNamespace ns, string cellName) {
            return ParseDouble(row.Elements(ns + "Cell")
                .FirstOrDefault(cell => string.Equals(cell.Attribute("N")?.Value, cellName, StringComparison.OrdinalIgnoreCase))
                ?.Attribute("V")?.Value);
        }

        private static bool TryGetTruthyCellValue(XElement element, string cellName) {
            string? value = element.Elements()
                .FirstOrDefault(child => string.Equals(child.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                                         string.Equals(child.Attribute("N")?.Value, cellName, StringComparison.OrdinalIgnoreCase))
                ?.Attribute("V")?.Value;
            return TryParseTruthyCellValue(value);
        }

        private static bool TryParseTruthyCellValue(string? value) {
            string? normalized = NormalizeCellLiteral(value);
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            if (bool.TryParse(normalized, out bool boolValue)) {
                return boolValue;
            }

            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue) &&
                   numericValue != 0;
        }

        private static bool TryParseCellIntValue(string? value, out int result) {
            string? normalized = NormalizeCellLiteral(value);
            if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out result)) {
                return true;
            }

            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue)) {
                int integerValue = Convert.ToInt32(numericValue);
                if (Math.Abs(numericValue - integerValue) <= 1e-9) {
                    result = integerValue;
                    return true;
                }
            }

            result = 0;
            return false;
        }

        private static string? NormalizeCellLiteral(string? value) {
            if (value is null) {
                return null;
            }

            string normalized = value.Trim();
            if (normalized.Length == 0) {
                return null;
            }
            while (normalized.StartsWith("GUARD(", StringComparison.OrdinalIgnoreCase) && normalized.EndsWith(")", StringComparison.Ordinal)) {
                normalized = normalized.Substring(6, normalized.Length - 7).Trim();
            }

            return normalized;
        }

        private static void CaptureConnectorShapeChildOrder(VisioConnector connector, XElement connectorElement) {
            connector.PreservedShapeChildren.Clear();
            foreach (XElement child in connectorElement.Elements()) {
                string localName = child.Name.LocalName;
                if (string.Equals(localName, "XForm1D", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(localName, "XForm", StringComparison.OrdinalIgnoreCase)) {
                    connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("XForm1D"));
                    continue;
                }

                if (string.Equals(localName, "Cell", StringComparison.OrdinalIgnoreCase)) {
                    string? cellName = child.Attribute("N")?.Value;
                    if (IsModeledConnectorCell(cellName)) {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry($"Cell:{cellName}"));
                    } else {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry(child));
                    }

                    continue;
                }

                if (string.Equals(localName, "Section", StringComparison.OrdinalIgnoreCase)) {
                    string? sectionName = child.Attribute("N")?.Value;
                    if (string.Equals(sectionName, "Geometry", StringComparison.OrdinalIgnoreCase)) {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("Section:Geometry"));
                    } else if (string.Equals(sectionName, "Hyperlink", StringComparison.OrdinalIgnoreCase)) {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("Section:Hyperlink"));
                    } else if (connector.HasModeledCharSection &&
                               IsCharacterSection(sectionName)) {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("Section:Char"));
                    } else if (connector.HasModeledParaSection &&
                               IsParagraphSection(sectionName)) {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("Section:Para"));
                    } else if (string.Equals(sectionName, "Prop", StringComparison.OrdinalIgnoreCase)) {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("Section:Prop"));
                    } else {
                        connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry(child));
                    }

                    continue;
                }

                if (string.Equals(localName, "Text", StringComparison.OrdinalIgnoreCase)) {
                    connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry("Text"));
                    continue;
                }

                connector.PreservedShapeChildren.Add(new VisioConnector.PreservedShapeChildEntry(child));
            }
        }

        private static bool IsModeledConnectorCell(string? cellName) {
            return string.Equals(cellName, "BeginX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "BeginY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "EndX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "EndY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LineWeight", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LinePattern", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LineColor", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "FillPattern", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "FillForegnd", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "OneD", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LayerMember", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "BeginArrow", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "EndArrow", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LeftMargin", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "RightMargin", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TopMargin", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "BottomMargin", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "VerticalAlign", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TextBkgnd", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TextBkgndTrans", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TxtPinX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TxtPinY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TxtWidth", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TxtHeight", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TxtLocPinX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "TxtLocPinY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapeRouteStyle", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ConLineRouteExt", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ConLineJumpStyle", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ConLineJumpCode", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ConLineJumpDirX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ConLineJumpDirY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ConFixedCode", StringComparison.OrdinalIgnoreCase) ||
                   VisioProtection.IsCellName(cellName);
        }

        private static bool ShouldPreserveConnectorCell(string? cellName) {
            return !string.IsNullOrWhiteSpace(cellName) &&
                   !string.Equals(cellName, "BeginArrow", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "EndArrow", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineWeight", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LinePattern", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineColor", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "FillPattern", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "FillForegnd", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "OneD", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LayerMember", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "BeginX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "BeginY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "EndX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "EndY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LeftMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "RightMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TopMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "BottomMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "VerticalAlign", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TextBkgnd", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TextBkgndTrans", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TxtPinX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TxtPinY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TxtWidth", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TxtHeight", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TxtLocPinX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "TxtLocPinY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapeRouteStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConLineRouteExt", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConLineJumpStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConLineJumpCode", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConLineJumpDirX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConLineJumpDirY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConFixedCode", StringComparison.OrdinalIgnoreCase) &&
                   !VisioProtection.IsCellName(cellName);
        }

        private static bool ShouldPreserveConnectorSection(VisioConnector connector, XElement section) {
            string? sectionName = section.Attribute("N")?.Value;
            if (connector.HasModeledCharSection &&
                IsCharacterSection(sectionName)) {
                return false;
            }

            if (connector.HasModeledParaSection &&
                IsParagraphSection(sectionName)) {
                return false;
            }

            return ShouldPreserveConnectorSection(section);
        }

        private static bool ShouldPreserveConnectorSection(XElement section) {
            string? sectionName = section.Attribute("N")?.Value;
            return !string.Equals(sectionName, "Geometry", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(sectionName, "Hyperlink", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(sectionName, "Prop", StringComparison.OrdinalIgnoreCase);
        }
    }
}
