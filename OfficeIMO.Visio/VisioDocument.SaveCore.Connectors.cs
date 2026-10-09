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
            VisioMaster? effectiveMaster = ResolveEffectiveMaster(connector);
            bool useDynamicMaster = effectiveMaster != null;
            string connName = useDynamicMaster ? "Dynamic connector" : "Connector";
            writer.WriteAttributeString("Name", connName);
            writer.WriteAttributeString("NameU", connName);
            writer.WriteAttributeString("Type", "Shape");
            if (useDynamicMaster) {
                writer.WriteAttributeString("Master", GetPackageMasterId(packageMasters, effectiveMaster!));
            }
            string? defaultStyle = useDynamicMaster ? null : "0";
            VisioNativeStyleReferences.Write(writer, connector.NativeStyleReferences, defaultStyle, defaultStyle, defaultStyle);

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
                        string? name = (string?)entry.RawElement.Attribute("N");
                        if (entry.RawElement.Name.LocalName == "Cell" && VisioConnectorGeometry.IsTransformCell(name)) {
                            WriteConnectorTransform(writer, ns, connector, name, entry.RawElement);
                            emittedTokens.Add("Cell:" + name);
                        } else entry.RawElement.WriteTo(writer);
                        continue;
                    }

                    if (entry.Token is string token &&
                        !string.IsNullOrWhiteSpace(token) &&
                        emittedTokens.Add(token) &&
                        TryWriteConnectorShapeChildToken(writer, ns, connector, connectorOriginalId, token, startX, startY, endX, endY, layerIndexes)) {
                        continue;
                    }
                }

                foreach (string name in ConnectorTransformCells)
                    if (emittedTokens.Add("Cell:" + name)) WriteConnectorTransform(writer, ns, connector, name);
                WriteRemainingConnectorShapeChildren(writer, ns, connector, connectorOriginalId, emittedTokens, startX, startY, endX, endY, layerIndexes);
                return;
            }

            WriteConnectorTransform(writer, ns, connector);
            WriteXForm1D(writer, ns, startX, startY, endX, endY);
            WriteModeledConnectorCells(writer, ns, connector, startX, startY, endX, endY, layerIndexes);
            WritePreservedConnectorCells(writer, connector.PreservedCellElements.Where(cell => !VisioConnectorGeometry.IsTransformCell((string?)cell.Attribute("N"))));
            WritePreservedConnectorSections(writer, connector.PreservedNonGeometrySections);
            WriteTextStyleSections(writer, ns, connector.TextStyle, connector.CharacterSectionSource, connector.ParagraphSectionSource, connector.PreservedNonGeometrySections);
            WriteHyperlinkSection(writer, ns, connector.Hyperlinks, VisioHyperlinkRowNames.Inherited(connector, this));
            WriteConnectorGeometry(writer, ns, connector, startX, startY, endX, endY);
            WriteDataSection(writer, ns, GetConnectorLabelData(connector), connector.PreservedDataRows, connectorOriginalId, connector.ShapeData, connector.ShapeDataSectionName);
            WriteTextElement(writer, ns, connector.Label, connector.PreservedTextElement, connector.PreservedTextValue);
        }

        private static void ComputeConnectorEndpoints(VisioConnector connector, out double startX, out double startY, out double endX, out double endY) {
            VisioConnectorEndpoints.Resolve(connector, out startX, out startY, out endX, out endY);
        }

        private bool TryWriteConnectorShapeChildToken(
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
                WriteHyperlinkSection(writer, ns, connector.Hyperlinks, VisioHyperlinkRowNames.Inherited(connector, this));
                return true;
            }

            if (string.Equals(token, "Section:Char", StringComparison.OrdinalIgnoreCase)) {
                WriteCharSection(writer, ns, connector.TextStyle, connector.CharacterSectionSource, connector.PreservedNonGeometrySections);
                return true;
            }

            if (string.Equals(token, "Section:Para", StringComparison.OrdinalIgnoreCase)) {
                WriteParaSection(writer, ns, connector.TextStyle, connector.ParagraphSectionSource, connector.PreservedNonGeometrySections);
                return true;
            }

            if (string.Equals(token, "Section:Prop", StringComparison.OrdinalIgnoreCase)) {
                WriteDataSection(writer, ns, GetConnectorLabelData(connector), connector.PreservedDataRows, connectorOriginalId, connector.ShapeData, connector.ShapeDataSectionName);
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

        private void WriteRemainingConnectorShapeChildren(
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
            bool hasEndpointCells = emittedTokens.Contains("Cell:BeginX") &&
                emittedTokens.Contains("Cell:BeginY") &&
                emittedTokens.Contains("Cell:EndX") &&
                emittedTokens.Contains("Cell:EndY");
            if (!hasEndpointCells && emittedTokens.Add("XForm1D")) {
                WriteXForm1D(writer, ns, startX, startY, endX, endY);
            }

            WriteRemainingModeledConnectorCells(writer, ns, connector, emittedTokens, startX, startY, endX, endY, layerIndexes);

            if (emittedTokens.Add("Section:Hyperlink")) {
                WriteHyperlinkSection(writer, ns, connector.Hyperlinks, VisioHyperlinkRowNames.Inherited(connector, this));
            }

            if (emittedTokens.Add("Section:Char")) {
                WriteCharSection(writer, ns, connector.TextStyle, connector.CharacterSectionSource, connector.PreservedNonGeometrySections);
            }

            if (emittedTokens.Add("Section:Para")) {
                WriteParaSection(writer, ns, connector.TextStyle, connector.ParagraphSectionSource, connector.PreservedNonGeometrySections);
            }

            if (emittedTokens.Add("Section:Geometry")) {
                WriteConnectorGeometry(writer, ns, connector, startX, startY, endX, endY);
            }

            if (emittedTokens.Add("Section:Prop")) {
                WriteDataSection(writer, ns, GetConnectorLabelData(connector), connector.PreservedDataRows, connectorOriginalId, connector.ShapeData, connector.ShapeDataSectionName);
            }

            if (emittedTokens.Add("Text")) {
                WriteTextElement(writer, ns, connector.Label, connector.PreservedTextElement, connector.PreservedTextValue);
            }
        }

        private void WriteModeledConnectorCells(XmlWriter writer, string ns, VisioConnector connector, double startX, double startY, double endX, double endY, IReadOnlyDictionary<string, int> layerIndexes) {
            foreach (string cellName in ConnectorModeledCellOrder) {
                TryWriteModeledConnectorCell(writer, ns, connector, cellName, startX, startY, endX, endY, layerIndexes);
            }
        }

        private void WriteRemainingModeledConnectorCells(XmlWriter writer, string ns, VisioConnector connector, ISet<string> emittedTokens, double startX, double startY, double endX, double endY, IReadOnlyDictionary<string, int> layerIndexes) {
            foreach (string cellName in ConnectorModeledCellOrder) {
                string token = GetModeledCellToken(cellName);
                if (emittedTokens.Add(token)) {
                    TryWriteModeledConnectorCell(writer, ns, connector, cellName, startX, startY, endX, endY, layerIndexes);
                }
            }
        }

        private bool TryWriteModeledConnectorCell(XmlWriter writer, string ns, VisioConnector connector, string cellName, double startX, double startY, double endX, double endY, IReadOnlyDictionary<string, int> layerIndexes) {
            switch (cellName) {
                case "BeginX":
                    WriteConnectorEndpointCell(writer, ns, connector, "BeginX", startX);
                    return true;
                case "BeginY":
                    WriteConnectorEndpointCell(writer, ns, connector, "BeginY", startY);
                    return true;
                case "EndX":
                    WriteConnectorEndpointCell(writer, ns, connector, "EndX", endX);
                    return true;
                case "EndY":
                    WriteConnectorEndpointCell(writer, ns, connector, "EndY", endY);
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
                    WriteLayerMemberCell(writer, ns, connector.LayerNames, layerIndexes, connector.NativeLayerMembership);
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
                    WriteTextBackgroundColorCell(writer, ns, connector.TextStyle);
                    return true;
                case "TextBkgndTrans":
                    WriteTextBackgroundTransparencyCell(writer, ns, connector.TextStyle);
                    return true;
                case "TxtAngle":
                    if (!string.IsNullOrEmpty(connector.Label) || connector.TextStyle?.TextAngle.HasValue == true || connector.LabelPlacement != null)
                        WriteCell(writer, ns, "TxtAngle", VisioConnectorLabelFrame.ResolveAngle(connector) - VisioConnectorGeometry.CreateShape(connector).Angle);
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

        private static void WriteConnectorEndpointCell(
            XmlWriter writer,
            string ns,
            VisioConnector connector,
            string cellName,
            double value) {
            if (connector.PreservedEndpointCellElements.TryGetValue(cellName,
                    out XElement? preserved)) {
                XElement current = new(preserved);
                current.Name = XName.Get("Cell", ns);
                current.SetAttributeValue("N", cellName);
                current.SetAttributeValue("V", ToVisioString(value));
                current.WriteTo(writer);
                return;
            }

            WriteCell(writer, ns, cellName, value);
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

            VisioConnectorLabelPlacement? placement = VisioConnectorLabelFrame.ResolvePlacement(connector);
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
            } else {
                (pinX, pinY) = ResolveConnectorPathPoint(connector, startX, startY, endX, endY, placement.Position);
                pinX += placement.OffsetX;
                pinY += placement.OffsetY;
            }
            VisioShape frame = VisioConnectorGeometry.CreateShape(connector);
            OfficeIMO.Drawing.OfficePoint local = VisioConnectorEndpoints.ToLocalPoint(frame, new OfficeIMO.Drawing.OfficePoint(pinX, pinY));
            bool native = connector.NativeGeometry?.AppliesTo(connector) == true;
            pinX = native && connector.NativeGeometry!.FlipX ? frame.Width - local.X : local.X;
            pinY = native && connector.NativeGeometry!.FlipY ? frame.Height - local.Y : local.Y;
            return true;
        }

        private static (double X, double Y) ResolveConnectorPathPoint(VisioConnector connector, double startX, double startY, double endX, double endY, double position) {
            List<(double X, double Y)> points = VisioConnectorGeometry.GetPoints(connector);

            return OfficeIMO.Drawing.OfficeGeometry.InterpolatePolyline(points, position);
        }
    }
}
