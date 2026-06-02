using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private void WriteConnectorShapeElement(XmlWriter writer, string ns, VisioConnector connector, IReadOnlyDictionary<string, string> persistedIds, IReadOnlyList<PackageMasterEntry> packageMasters, IReadOnlyDictionary<string, int> layerIndexes) {
            writer.WriteStartElement("Shape", ns);
            writer.WriteAttributeString("ID", GetPersistedId(persistedIds, connector.Id));
            bool isDynamic = connector.Kind == ConnectorKind.Dynamic;
            string connName = (isDynamic && UseMastersByDefault) ? "Dynamic connector" : "Connector";
            writer.WriteAttributeString("Name", connName);
            writer.WriteAttributeString("NameU", connName);
            writer.WriteAttributeString("LineStyle", "0");
            writer.WriteAttributeString("FillStyle", "0");
            writer.WriteAttributeString("TextStyle", "0");
            if (isDynamic && UseMastersByDefault) {
                var m = EnsureBuiltinMaster("Dynamic connector");
                writer.WriteAttributeString("Master", GetPackageMasterId(packageMasters, m));
            }

            WriteConnectorShapeBody(writer, ns, connector, persistedIds, layerIndexes);
            writer.WriteEndElement();
        }

        private void WriteConnectorShapeBody(XmlWriter writer, string ns, VisioConnector connector, IReadOnlyDictionary<string, string> persistedIds, IReadOnlyDictionary<string, int> layerIndexes) {
            ComputeConnectorEndpoints(connector, out double startX, out double startY, out double endX, out double endY);
            KeyValuePair<string, string>? connectorOriginalId = GetOriginalIdEntry(persistedIds, connector.Id);

            if (connector.PreservedShapeChildren.Count > 0) {
                HashSet<string> emittedTokens = new(StringComparer.OrdinalIgnoreCase);
                foreach (VisioConnector.PreservedShapeChildEntry entry in connector.PreservedShapeChildren) {
                    if (entry.RawElement != null) {
                        entry.RawElement.WriteTo(writer);
                        continue;
                    }

                    if (entry.Token is string token &&
                        !string.IsNullOrWhiteSpace(token) &&
                        emittedTokens.Add(token) &&
                        TryWriteConnectorShapeChildToken(writer, ns, connector, connectorOriginalId, token, startX, startY, endX, endY, layerIndexes)) {
                        continue;
                    }
                }

                WriteRemainingConnectorShapeChildren(writer, ns, connector, connectorOriginalId, emittedTokens, startX, startY, endX, endY, layerIndexes);
                return;
            }

            WriteXForm1D(writer, ns, startX, startY, endX, endY);
            WriteModeledConnectorCells(writer, ns, connector, startX, startY, endX, endY, layerIndexes);
            WritePreservedConnectorCells(writer, connector.PreservedCellElements);
            WritePreservedConnectorSections(writer, connector.PreservedNonGeometrySections);
            WriteTextStyleSections(writer, ns, connector.TextStyle);
            WriteHyperlinkSection(writer, ns, connector.Hyperlinks);
            WriteConnectorGeometry(writer, ns, connector, startX, startY, endX, endY);
            WriteDataSection(writer, ns, connector.Data, connector.PreservedDataRows, connectorOriginalId, connector.ShapeData);
            WriteTextElement(writer, ns, connector.Label, connector.PreservedTextElement, connector.PreservedTextValue);
        }

        private static void ComputeConnectorEndpoints(VisioConnector connector, out double startX, out double startY, out double endX, out double endY) {
            if (connector.FromConnectionPoint != null) {
                (startX, startY) = connector.From.GetAbsolutePoint(connector.FromConnectionPoint.X, connector.FromConnectionPoint.Y);
            } else {
                var (fL, fB, fR, fT) = connector.From.GetBounds();
                var (tL2, _, tR2, _) = connector.To.GetBounds();
                double fromCx = (fL + fR) / 2.0;
                double toCx = (tL2 + tR2) / 2.0;
                bool toIsRight = toCx >= fromCx;
                startX = toIsRight ? fR : fL;
                startY = (fB + fT) / 2.0;
            }

            if (connector.ToConnectionPoint != null) {
                (endX, endY) = connector.To.GetAbsolutePoint(connector.ToConnectionPoint.X, connector.ToConnectionPoint.Y);
            } else {
                var (tL, tB, tR, tT) = connector.To.GetBounds();
                var (fL2, _, fR2, _) = connector.From.GetBounds();
                double toCx = (tL + tR) / 2.0;
                double fromCx = (fL2 + fR2) / 2.0;
                bool fromIsLeft = fromCx <= toCx;
                endX = fromIsLeft ? tL : tR;
                endY = (tB + tT) / 2.0;
            }
        }

        private static bool TryWriteConnectorShapeChildToken(
            XmlWriter writer,
            string ns,
            VisioConnector connector,
            KeyValuePair<string, string>? connectorOriginalId,
            string token,
            double startX,
            double startY,
            double endX,
            double endY,
            IReadOnlyDictionary<string, int> layerIndexes) {
            if (string.Equals(token, "XForm1D", StringComparison.OrdinalIgnoreCase)) {
                WriteXForm1D(writer, ns, startX, startY, endX, endY);
                return true;
            }

            if (string.Equals(token, "Section:Geometry", StringComparison.OrdinalIgnoreCase)) {
                WriteConnectorGeometry(writer, ns, connector, startX, startY, endX, endY);
                return true;
            }

            if (string.Equals(token, "Section:Hyperlink", StringComparison.OrdinalIgnoreCase)) {
                WriteHyperlinkSection(writer, ns, connector.Hyperlinks);
                return true;
            }

            if (string.Equals(token, "Section:Char", StringComparison.OrdinalIgnoreCase)) {
                WriteCharSection(writer, ns, connector.TextStyle);
                return true;
            }

            if (string.Equals(token, "Section:Para", StringComparison.OrdinalIgnoreCase)) {
                WriteParaSection(writer, ns, connector.TextStyle);
                return true;
            }

            if (string.Equals(token, "Section:Prop", StringComparison.OrdinalIgnoreCase)) {
                WriteDataSection(writer, ns, connector.Data, connector.PreservedDataRows, connectorOriginalId, connector.ShapeData);
                return true;
            }

            if (string.Equals(token, "Text", StringComparison.OrdinalIgnoreCase)) {
                WriteTextElement(writer, ns, connector.Label, connector.PreservedTextElement, connector.PreservedTextValue);
                return true;
            }

            if (token.StartsWith("Cell:", StringComparison.OrdinalIgnoreCase)) {
                return TryWriteModeledConnectorCell(writer, ns, connector, token.Substring("Cell:".Length), startX, startY, endX, endY, layerIndexes);
            }

            return false;
        }

        private static void WriteRemainingConnectorShapeChildren(
            XmlWriter writer,
            string ns,
            VisioConnector connector,
            KeyValuePair<string, string>? connectorOriginalId,
            ISet<string> emittedTokens,
            double startX,
            double startY,
            double endX,
            double endY,
            IReadOnlyDictionary<string, int> layerIndexes) {
            if (emittedTokens.Add("XForm1D")) {
                WriteXForm1D(writer, ns, startX, startY, endX, endY);
            }

            WriteRemainingModeledConnectorCells(writer, ns, connector, emittedTokens, startX, startY, endX, endY, layerIndexes);

            if (emittedTokens.Add("Section:Hyperlink")) {
                WriteHyperlinkSection(writer, ns, connector.Hyperlinks);
            }

            if (emittedTokens.Add("Section:Char")) {
                WriteCharSection(writer, ns, connector.TextStyle);
            }

            if (emittedTokens.Add("Section:Para")) {
                WriteParaSection(writer, ns, connector.TextStyle);
            }

            if (emittedTokens.Add("Section:Geometry")) {
                WriteConnectorGeometry(writer, ns, connector, startX, startY, endX, endY);
            }

            if (emittedTokens.Add("Section:Prop")) {
                WriteDataSection(writer, ns, connector.Data, connector.PreservedDataRows, connectorOriginalId, connector.ShapeData);
            }

            if (emittedTokens.Add("Text")) {
                WriteTextElement(writer, ns, connector.Label, connector.PreservedTextElement, connector.PreservedTextValue);
            }
        }

        private static void WriteModeledConnectorCells(XmlWriter writer, string ns, VisioConnector connector, double startX, double startY, double endX, double endY, IReadOnlyDictionary<string, int> layerIndexes) {
            foreach (string cellName in ConnectorModeledCellOrder) {
                TryWriteModeledConnectorCell(writer, ns, connector, cellName, startX, startY, endX, endY, layerIndexes);
            }
        }

        private static void WriteRemainingModeledConnectorCells(XmlWriter writer, string ns, VisioConnector connector, ISet<string> emittedTokens, double startX, double startY, double endX, double endY, IReadOnlyDictionary<string, int> layerIndexes) {
            foreach (string cellName in ConnectorModeledCellOrder) {
                string token = $"Cell:{cellName}";
                if (emittedTokens.Add(token)) {
                    TryWriteModeledConnectorCell(writer, ns, connector, cellName, startX, startY, endX, endY, layerIndexes);
                }
            }
        }

        private static bool TryWriteModeledConnectorCell(XmlWriter writer, string ns, VisioConnector connector, string cellName, double startX, double startY, double endX, double endY, IReadOnlyDictionary<string, int> layerIndexes) {
            switch (cellName) {
                case "BeginX":
                    WriteCell(writer, ns, "BeginX", startX);
                    return true;
                case "BeginY":
                    WriteCell(writer, ns, "BeginY", startY);
                    return true;
                case "EndX":
                    WriteCell(writer, ns, "EndX", endX);
                    return true;
                case "EndY":
                    WriteCell(writer, ns, "EndY", endY);
                    return true;
                case "LineWeight":
                    WriteCell(writer, ns, "LineWeight", connector.LineWeight);
                    return true;
                case "LinePattern":
                    WriteCell(writer, ns, "LinePattern", connector.LinePattern);
                    return true;
                case "LineColor":
                    WriteCellValue(writer, ns, "LineColor", connector.LineColor.ToVisioHex());
                    return true;
                case "FillPattern":
                    WriteCell(writer, ns, "FillPattern", 0);
                    return true;
                case "FillForegnd":
                    WriteCellValue(writer, ns, "FillForegnd", Color.Transparent.ToVisioHex());
                    return true;
                case "OneD":
                    WriteCell(writer, ns, "OneD", 1);
                    return true;
                case "LayerMember":
                    WriteLayerMemberCell(writer, ns, connector.LayerNames, layerIndexes);
                    return true;
                case "ShapeRouteStyle":
                    if (connector.RouteStyle.HasValue) {
                        WriteCell(writer, ns, "ShapeRouteStyle", (int)connector.RouteStyle.Value);
                    }
                    return true;
                case "ConLineRouteExt":
                    if (connector.RouteAppearance.HasValue) {
                        WriteCell(writer, ns, "ConLineRouteExt", (int)connector.RouteAppearance.Value);
                    }
                    return true;
                case "ConLineJumpStyle":
                    if (connector.LineJumpStyle.HasValue) {
                        WriteCell(writer, ns, "ConLineJumpStyle", (int)connector.LineJumpStyle.Value);
                    }
                    return true;
                case "ConLineJumpCode":
                    if (connector.LineJumpCode.HasValue) {
                        WriteCell(writer, ns, "ConLineJumpCode", (int)connector.LineJumpCode.Value);
                    }
                    return true;
                case "ConLineJumpDirX":
                    if (connector.HorizontalJumpDirection.HasValue) {
                        WriteCell(writer, ns, "ConLineJumpDirX", (int)connector.HorizontalJumpDirection.Value);
                    }
                    return true;
                case "ConLineJumpDirY":
                    if (connector.VerticalJumpDirection.HasValue) {
                        WriteCell(writer, ns, "ConLineJumpDirY", (int)connector.VerticalJumpDirection.Value);
                    }
                    return true;
                case "ConFixedCode":
                    if (connector.RerouteBehavior.HasValue) {
                        WriteCell(writer, ns, "ConFixedCode", (int)connector.RerouteBehavior.Value);
                    }
                    return true;
                case "BeginArrow":
                    if (connector.BeginArrow.HasValue) {
                        WriteCell(writer, ns, "BeginArrow", (int)connector.BeginArrow.Value);
                    }
                    return true;
                case "EndArrow":
                    if (connector.EndArrow.HasValue) {
                        WriteCell(writer, ns, "EndArrow", (int)connector.EndArrow.Value);
                    }
                    return true;
                case "LeftMargin":
                    if (connector.TextStyle?.LeftMargin.HasValue == true) {
                        WriteCell(writer, ns, "LeftMargin", connector.TextStyle.LeftMargin.Value);
                    }
                    return true;
                case "RightMargin":
                    if (connector.TextStyle?.RightMargin.HasValue == true) {
                        WriteCell(writer, ns, "RightMargin", connector.TextStyle.RightMargin.Value);
                    }
                    return true;
                case "TopMargin":
                    if (connector.TextStyle?.TopMargin.HasValue == true) {
                        WriteCell(writer, ns, "TopMargin", connector.TextStyle.TopMargin.Value);
                    }
                    return true;
                case "BottomMargin":
                    if (connector.TextStyle?.BottomMargin.HasValue == true) {
                        WriteCell(writer, ns, "BottomMargin", connector.TextStyle.BottomMargin.Value);
                    }
                    return true;
                case "VerticalAlign":
                    if (connector.TextStyle?.VerticalAlignment.HasValue == true) {
                        WriteCell(writer, ns, "VerticalAlign", (int)connector.TextStyle.VerticalAlignment.Value);
                    }
                    return true;
                case "TextBkgnd":
                    if (connector.TextStyle?.BackgroundColor.HasValue == true) {
                        WriteCellValue(writer, ns, "TextBkgnd", connector.TextStyle.BackgroundColor.Value.ToVisioHex());
                    }
                    return true;
                case "TextBkgndTrans":
                    if (connector.TextStyle?.BackgroundTransparency.HasValue == true) {
                        WriteCell(writer, ns, "TextBkgndTrans", connector.TextStyle.BackgroundTransparency.Value);
                    }
                    return true;
                case "TxtPinX":
                    if (TryResolveConnectorLabelPlacement(connector, startX, startY, endX, endY, out double txtPinX, out _, out _, out _, out _, out _)) {
                        WriteCell(writer, ns, "TxtPinX", txtPinX);
                    }
                    return true;
                case "TxtPinY":
                    if (TryResolveConnectorLabelPlacement(connector, startX, startY, endX, endY, out _, out double txtPinY, out _, out _, out _, out _)) {
                        WriteCell(writer, ns, "TxtPinY", txtPinY);
                    }
                    return true;
                case "TxtWidth":
                    if (TryResolveConnectorLabelPlacement(connector, startX, startY, endX, endY, out _, out _, out double txtWidth, out _, out _, out _)) {
                        WriteCell(writer, ns, "TxtWidth", txtWidth);
                    }
                    return true;
                case "TxtHeight":
                    if (TryResolveConnectorLabelPlacement(connector, startX, startY, endX, endY, out _, out _, out _, out double txtHeight, out _, out _)) {
                        WriteCell(writer, ns, "TxtHeight", txtHeight);
                    }
                    return true;
                case "TxtLocPinX":
                    if (TryResolveConnectorLabelPlacement(connector, startX, startY, endX, endY, out _, out _, out _, out _, out double txtLocPinX, out _)) {
                        WriteCell(writer, ns, "TxtLocPinX", txtLocPinX);
                    }
                    return true;
                case "TxtLocPinY":
                    if (TryResolveConnectorLabelPlacement(connector, startX, startY, endX, endY, out _, out _, out _, out _, out _, out double txtLocPinY)) {
                        WriteCell(writer, ns, "TxtLocPinY", txtLocPinY);
                    }
                    return true;
                default:
                    if (TryWriteProtectionCell(writer, ns, connector.Protection, cellName)) {
                        return true;
                    }

                    return false;
            }
        }

        private static bool TryResolveConnectorLabelPlacement(
            VisioConnector connector,
            double startX,
            double startY,
            double endX,
            double endY,
            out double pinX,
            out double pinY,
            out double width,
            out double height,
            out double locPinX,
            out double locPinY) {
            pinX = 0D;
            pinY = 0D;
            width = 0D;
            height = 0D;
            locPinX = 0D;
            locPinY = 0D;

            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            if (placement == null) {
                return false;
            }

            width = placement.Width;
            height = placement.Height;
            locPinX = placement.GetLocPinX();
            locPinY = placement.GetLocPinY();

            if (placement.AbsolutePinX.HasValue && placement.AbsolutePinY.HasValue) {
                pinX = placement.AbsolutePinX.Value;
                pinY = placement.AbsolutePinY.Value;
                return true;
            }

            (pinX, pinY) = ResolveConnectorPathPoint(connector, startX, startY, endX, endY, placement.Position);
            pinX += placement.OffsetX;
            pinY += placement.OffsetY;
            return true;
        }

        private static (double X, double Y) ResolveConnectorPathPoint(VisioConnector connector, double startX, double startY, double endX, double endY, double position) {
            double clampedPosition = VisioConnectorLabelPlacement.ClampPosition(position);
            List<(double X, double Y)> points = new() {
                (startX, startY)
            };

            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    points.Add((waypoint.X, waypoint.Y));
                }
            } else if (connector.Kind == ConnectorKind.RightAngle) {
                points.Add((startX, endY));
            }

            points.Add((endX, endY));

            double totalLength = 0D;
            for (int i = 1; i < points.Count; i++) {
                totalLength += Distance(points[i - 1], points[i]);
            }

            if (totalLength <= 0D) {
                return (startX, startY);
            }

            double targetLength = totalLength * clampedPosition;
            double traversed = 0D;
            for (int i = 1; i < points.Count; i++) {
                (double X, double Y) from = points[i - 1];
                (double X, double Y) to = points[i];
                double segmentLength = Distance(from, to);
                if (segmentLength <= 0D) {
                    continue;
                }

                if (traversed + segmentLength >= targetLength) {
                    double segmentPosition = (targetLength - traversed) / segmentLength;
                    return (
                        from.X + ((to.X - from.X) * segmentPosition),
                        from.Y + ((to.Y - from.Y) * segmentPosition));
                }

                traversed += segmentLength;
            }

            return (endX, endY);
        }

        private static double Distance((double X, double Y) from, (double X, double Y) to) {
            double dx = to.X - from.X;
            double dy = to.Y - from.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }
    }
}
