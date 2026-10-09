using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private static VisioConnector? LoadConnector(XElement connectorElement, XNamespace vNs,
        IReadOnlyDictionary<string, VisioMaster> masters, IReadOnlyDictionary<int, string> faceNamesById,
        VisioPage page, VisioShape? fromShape, VisioShape? toShape, string? fromCell, string? toCell, IReadOnlyDictionary<int, string>? textBackgroundColors) {
        OfficePoint? startPoint = ReadConnectorPoint(connectorElement, vNs, "Begin");
        OfficePoint? endPoint = ReadConnectorPoint(connectorElement, vNs, "End");
        if (fromShape == null && !startPoint.HasValue || toShape == null && !endPoint.HasValue) return null;
        string persistedId = (string?)connectorElement.Attribute("ID") ?? string.Empty;
        var connector = new VisioConnector(GetOriginalId(connectorElement, vNs) ?? persistedId,
            startPoint ?? default, endPoint ?? default) {
            From = fromShape, To = toShape, PersistedId = persistedId,
            NativeStyleReferences = VisioNativeStyleReferences.Read(connectorElement),
            PreserveDynamicConnectorMaster = HasDynamicConnectorIdentity(connectorElement, masters)
        };
        foreach (XElement cell in connectorElement.Elements(vNs + "Cell")) {
            string? n = cell.Attribute("N")?.Value;
            string? v = cell.Attribute("V")?.Value;
            switch (n) {
                case "BeginArrow":
                    if (TryParseCellIntValue(v, out int beginArrow)) {
                        connector.BeginArrow = (EndArrow)beginArrow;
                    }
                    break;
                case "EndArrow":
                    if (TryParseCellIntValue(v, out int endArrow)) {
                        connector.EndArrow = (EndArrow)endArrow;
                    }
                    break;
                case "LineWeight":
                    connector.LineWeight = ParseDouble(v);
                    break;
                case "LinePattern":
                    if (TryParseCellIntValue(v, out int connectorLinePattern)) {
                        connector.LinePattern = connectorLinePattern;
                    }
                    break;
                case "LineColor":
                    connector.LineColor = ParseColor(v, connector.LineColor);
                    break;
                case "LeftMargin":
                    EnsureConnectorTextStyle(connector).LeftMargin = ParseDouble(v);
                    break;
                case "RightMargin":
                    EnsureConnectorTextStyle(connector).RightMargin = ParseDouble(v);
                    break;
                case "TopMargin":
                    EnsureConnectorTextStyle(connector).TopMargin = ParseDouble(v);
                    break;
                case "BottomMargin":
                    EnsureConnectorTextStyle(connector).BottomMargin = ParseDouble(v);
                    break;
                case "VerticalAlign":
                    if (TryParseCellIntValue(v, out int connectorVerticalAlign) &&
                        Enum.IsDefined(typeof(VisioTextVerticalAlignment), connectorVerticalAlign)) {
                        EnsureConnectorTextStyle(connector).VerticalAlignment = (VisioTextVerticalAlignment)connectorVerticalAlign;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "TextBkgnd":
                    LoadTextBackgroundColor(EnsureConnectorTextStyle(connector), cell, textBackgroundColors);
                    break;
                case "TextBkgndTrans":
                    LoadTextBackgroundTransparency(EnsureConnectorTextStyle(connector), cell);
                    break;
                case "TxtAngle":
                    EnsureConnectorTextStyle(connector).TextAngle = ParseDouble(v);
                    break;
                case "TxtPinX":
                    EnsureConnectorLabelPlacement(connector).AbsolutePinX = ParseDouble(v);
                    break;
                case "TxtPinY":
                    EnsureConnectorLabelPlacement(connector).AbsolutePinY = ParseDouble(v);
                    break;
                case "TxtWidth":
                    EnsureConnectorLabelPlacement(connector).Width = ParseDouble(v);
                    break;
                case "TxtHeight":
                    EnsureConnectorLabelPlacement(connector).Height = ParseDouble(v);
                    break;
                case "TxtLocPinX":
                    EnsureConnectorLabelPlacement(connector).LocPinX = ParseDouble(v);
                    break;
                case "TxtLocPinY":
                    EnsureConnectorLabelPlacement(connector).LocPinY = ParseDouble(v);
                    break;
                case "LayerMember":
                    ParseLayerIndexes(v, connector.LayerIndexes);
                    connector.NativeLayerMembership = new VisioLayerMembership(cell, connector.LayerIndexes);
                    break;
                case "ShapeRouteStyle":
                    if (TryParseCellIntValue(v, out int connectorRouteStyle) &&
                        Enum.IsDefined(typeof(VisioPageRouteStyle), connectorRouteStyle)) {
                        connector.RouteStyle = (VisioPageRouteStyle)connectorRouteStyle;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "ConLineRouteExt":
                    if (TryParseCellIntValue(v, out int connectorRouteAppearance) &&
                        Enum.IsDefined(typeof(VisioLineRouteExtension), connectorRouteAppearance)) {
                        connector.RouteAppearance = (VisioLineRouteExtension)connectorRouteAppearance;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "ConLineJumpStyle":
                    if (TryParseCellIntValue(v, out int connectorJumpStyle) &&
                        Enum.IsDefined(typeof(VisioLineJumpStyle), connectorJumpStyle)) {
                        connector.LineJumpStyle = (VisioLineJumpStyle)connectorJumpStyle;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "ConLineJumpCode":
                    if (TryParseCellIntValue(v, out int connectorJumpCode) &&
                        Enum.IsDefined(typeof(VisioConnectorLineJumpCode), connectorJumpCode)) {
                        connector.LineJumpCode = (VisioConnectorLineJumpCode)connectorJumpCode;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "ConLineJumpDirX":
                    if (TryParseCellIntValue(v, out int connectorJumpDirX) &&
                        Enum.IsDefined(typeof(VisioHorizontalLineJumpDirection), connectorJumpDirX)) {
                        connector.HorizontalJumpDirection = (VisioHorizontalLineJumpDirection)connectorJumpDirX;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "ConLineJumpDirY":
                    if (TryParseCellIntValue(v, out int connectorJumpDirY) &&
                        Enum.IsDefined(typeof(VisioVerticalLineJumpDirection), connectorJumpDirY)) {
                        connector.VerticalJumpDirection = (VisioVerticalLineJumpDirection)connectorJumpDirY;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                case "ConFixedCode":
                    if (TryParseCellIntValue(v, out int connectorRerouteBehavior) &&
                        Enum.IsDefined(typeof(VisioConnectorRerouteBehavior), connectorRerouteBehavior)) {
                        connector.RerouteBehavior = (VisioConnectorRerouteBehavior)connectorRerouteBehavior;
                    } else {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
                default:
                    if (VisioProtection.IsCellName(n) &&
                        connector.Protection.TrySetCellValue(n, ParseNullableBoolCell(v))) {
                        break;
                    }

                    if (ShouldPreserveConnectorCell(n)) {
                        connector.PreservedCellElements.Add(new XElement(cell));
                    }
                    break;
            }
        }

        connector.Kind = DetermineConnectorKind(connectorElement, vNs, masters);
        ApplyLayerNamesFromIndexes(page, connector);
        XElement? connectorCharSection = connectorElement.Elements(vNs + "Section")
            .FirstOrDefault(section => IsCharacterSection(section.Attribute("N")?.Value));
        if (connectorCharSection != null && TryParseSimpleConnectorCharSection(connector, connectorCharSection, vNs, faceNamesById)) {
            connector.HasModeledCharSection = true;
            connector.CharacterSectionSource = CaptureTextSection(connectorCharSection, connector.TextStyle, character: true);
        }

        XElement? connectorParaSection = connectorElement.Elements(vNs + "Section")
            .FirstOrDefault(section => IsParagraphSection(section.Attribute("N")?.Value));
        if (connectorParaSection != null && TryParseSimpleConnectorParaSection(connector, connectorParaSection, vNs)) {
            connector.HasModeledParaSection = true;
            connector.ParagraphSectionSource = CaptureTextSection(connectorParaSection, connector.TextStyle, character: false);
        }

        foreach (XElement geometrySection in connectorElement.Elements(vNs + "Section")
                     .Where(section => string.Equals(section.Attribute("N")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))) {
            connector.PreservedGeometrySections.Add(new XElement(geometrySection));
        }
        foreach (XElement section in connectorElement.Elements(vNs + "Section")
                     .Where(section => ShouldPreserveConnectorSection(connector, section))) {
            connector.PreservedNonGeometrySections.Add(new XElement(section));
        }
        XElement? connectorHyperlinkSection = connectorElement.Elements(vNs + "Section")
            .FirstOrDefault(section => string.Equals(section.Attribute("N")?.Value, "Hyperlink", StringComparison.OrdinalIgnoreCase));
        if (connectorHyperlinkSection != null) {
            ParseHyperlinks(connectorHyperlinkSection, vNs, connector.Hyperlinks);
        }

        XElement? connectorPropSection = connectorElement.Elements(vNs + "Section")
            .FirstOrDefault(section => IsShapeDataSectionName(section.Attribute("N")?.Value));
        if (connectorPropSection != null) {
            connector.ShapeDataSectionName = connectorPropSection.Attribute("N")?.Value ?? "Prop";
            ParseShapeDataRows(connectorPropSection, vNs, connector.ShapeData, connector.PreservedDataRows, connector.Data);
        }

        connector.FromConnectionPoint = fromShape == null ? null : ResolveConnectionPoint(fromShape, fromCell);
        connector.ToConnectionPoint = toShape == null ? null : ResolveConnectionPoint(toShape, toCell);
        if (fromShape != null && startPoint.HasValue) connector.StartAttachment = new VisioConnectorAttachment(fromShape, startPoint.Value);
        if (toShape != null && endPoint.HasValue) connector.EndAttachment = new VisioConnectorAttachment(toShape, endPoint.Value);
        VisioShape frame = new("connector-frame") {
            PinX = GetCellValue(connectorElement, vNs, "PinX"), PinY = GetCellValue(connectorElement, vNs, "PinY"),
            Width = GetCellValue(connectorElement, vNs, "Width"), Height = GetCellValue(connectorElement, vNs, "Height"),
            LocPinX = GetCellValue(connectorElement, vNs, "LocPinX"), LocPinY = GetCellValue(connectorElement, vNs, "LocPinY"),
            Angle = GetCellValue(connectorElement, vNs, "Angle")
        };
        connector.NativeGeometry = new VisioConnectorNativeGeometry(frame, startPoint ?? connector.StartPoint, endPoint ?? connector.EndPoint,
            connector.Kind, TryGetTruthyCellValue(connectorElement, "FlipX"), TryGetTruthyCellValue(connectorElement, "FlipY"));
        if (connector.LabelPlacement is VisioConnectorLabelPlacement placement && placement.PinX.HasValue && placement.PinY.HasValue) {
            (double x, double y) = connector.NativeGeometry.SourceToPage(placement.PinX.Value, placement.PinY.Value);
            placement.SetAbsolutePin(x, y);
        }
        if (connector.TextStyle?.TextAngle.HasValue == true || connector.LabelPlacement != null)
            EnsureConnectorTextStyle(connector).TextAngle = frame.Angle + (connector.TextStyle?.TextAngle ?? 0);
        bool restoredLabelAnchor = RestoreConnectorLabelAnchor(connector);
        if (connector.LabelPlacement is VisioConnectorLabelPlacement nativePlacement &&
            (!restoredLabelAnchor || nativePlacement.AnchorKind == VisioConnectorLabelAnchorKind.Native) &&
            nativePlacement.PinX.HasValue && nativePlacement.PinY.HasValue) {
            nativePlacement.AnchorKind = VisioConnectorLabelAnchorKind.Native;
            nativePlacement.NativeFrame = new VisioConnectorNativeTextFrame(nativePlacement.PinX.Value, nativePlacement.PinY.Value, connector.TextStyle?.TextAngle ?? 0);
        }
        TryHydrateConnectorWaypoints(connector, connectorElement, vNs);
        connector.NativeGeometry.CaptureRoute(connector);
        connector.PreservedFromConnectionCell = fromCell;
        connector.PreservedToConnectionCell = toCell;
        XElement? connectorTextElement = connectorElement.Element(vNs + "Text");
        connector.Label = connectorTextElement?.Value;
        connector.PreservedTextElement = connectorTextElement != null ? new XElement(connectorTextElement) : null;
        connector.PreservedTextValue = connectorTextElement?.Value;
        return connector;
    }

    private static OfficePoint? ReadConnectorPoint(XElement element, XNamespace ns, string prefix) {
        if (TryGetNumericCellValue(element, ns, prefix + "X", out double x) &&
            TryGetNumericCellValue(element, ns, prefix + "Y", out double y)) return new OfficePoint(x, y);
        return null;
    }
}
