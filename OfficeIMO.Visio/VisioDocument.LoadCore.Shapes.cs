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

        private static VisioShape ParseShape(XElement shapeElement, XNamespace ns, VisioShape? parent = null, int depth = 0) {
            return ParseShapeCore(shapeElement, ns, null, parent, depth);
        }

        private static VisioShape ParseShapeCore(XElement shapeElement, XNamespace ns, IReadOnlyDictionary<int, string>? faceNamesById = null, VisioShape? parent = null, int depth = 0) {
            if (depth > MaxShapeNestingDepth) {
                throw new InvalidOperationException("Maximum nesting depth exceeded");
            }

            VisioShape shape = ParseShapeBasics(shapeElement, ns);
            shape.Parent = parent;

            ParseShapeTransform(shape, shapeElement, ns);
            ParseShapeProperties(shape, shapeElement, ns, faceNamesById);
            ParseChildShapes(shape, shapeElement, ns, faceNamesById, depth);
            CaptureShapeChildOrder(shape, shapeElement);

            return shape;
        }

        private static VisioShape ParseShapeBasics(XElement shapeElement, XNamespace ns) {
            string persistedId = shapeElement.Attribute("ID")?.Value ?? string.Empty;
            string id = GetOriginalId(shapeElement, ns) ?? persistedId;
            XElement? textElement = shapeElement.Element(ns + "Text");
            VisioShape shape = new(id) {
                Name = shapeElement.Attribute("Name")?.Value,
                NameU = shapeElement.Attribute("NameU")?.Value,
                Text = textElement?.Value
            };

            shape.PersistedId = persistedId;
            shape.Type = shapeElement.Attribute("Type")?.Value;
            shape.MasterShapeId = shapeElement.Attribute("MasterShape")?.Value;
            shape.PreservedTextElement = textElement != null ? new XElement(textElement) : null;
            shape.PreservedTextValue = textElement?.Value;

            return shape;
        }

        private static void ParseShapeTransform(VisioShape shape, XElement shapeElement, XNamespace ns) {
            List<XElement> cellElements = shapeElement.Elements(ns + "Cell").ToList();
            bool pinXFound = false;
            bool pinYFound = false;
            bool widthFound = false;
            bool heightFound = false;
            bool locPinXFound = false;
            bool locPinYFound = false;
            bool angleFound = false;
            bool lineWeightFound = false;

            foreach (XElement cell in cellElements) {
                string? n = cell.Attribute("N")?.Value;
                string? v = cell.Attribute("V")?.Value;
                switch (n) {
                    case "PinX":
                        shape.PinX = ParseDouble(v);
                        pinXFound = true;
                        break;
                    case "PinY":
                        shape.PinY = ParseDouble(v);
                        pinYFound = true;
                        break;
                    case "Width":
                        shape.Width = ParseDouble(v);
                        widthFound = true;
                        shape.HasExplicitWidth = true;
                        break;
                    case "Height":
                        shape.Height = ParseDouble(v);
                        heightFound = true;
                        shape.HasExplicitHeight = true;
                        break;
                    case "LocPinX":
                        shape.LocPinX = ParseDouble(v);
                        locPinXFound = true;
                        shape.HasExplicitLocPinX = true;
                        break;
                    case "LocPinY":
                        shape.LocPinY = ParseDouble(v);
                        locPinYFound = true;
                        shape.HasExplicitLocPinY = true;
                        break;
                    case "Angle":
                        shape.Angle = ParseDouble(v);
                        angleFound = true;
                        break;
                    case "LineWeight":
                        shape.LineWeight = ParseDouble(v);
                        lineWeightFound = true;
                        break;
                    case "LinePattern":
                        if (TryParseCellIntValue(v, out int linePattern)) {
                            shape.LinePattern = linePattern;
                        }
                        break;
                    case "FillPattern":
                        if (TryParseCellIntValue(v, out int fillPattern)) {
                            shape.FillPattern = fillPattern;
                        }
                        break;
                    case "LineColor":
                        shape.LineColor = ParseColor(v, shape.LineColor);
                        break;
                    case "FillForegnd":
                        shape.FillColor = ParseColor(v, shape.FillColor);
                        break;
                    case "LeftMargin":
                        EnsureTextStyle(shape).LeftMargin = ParseDouble(v);
                        break;
                    case "RightMargin":
                        EnsureTextStyle(shape).RightMargin = ParseDouble(v);
                        break;
                    case "TopMargin":
                        EnsureTextStyle(shape).TopMargin = ParseDouble(v);
                        break;
                    case "BottomMargin":
                        EnsureTextStyle(shape).BottomMargin = ParseDouble(v);
                        break;
                    case "VerticalAlign":
                        if (TryParseCellIntValue(v, out int verticalAlign) &&
                            Enum.IsDefined(typeof(VisioTextVerticalAlignment), verticalAlign)) {
                            EnsureTextStyle(shape).VerticalAlignment = (VisioTextVerticalAlignment)verticalAlign;
                        } else {
                            shape.PreservedCellElements.Add(new XElement(cell));
                        }

                        break;
                    case "TextBkgnd":
                        EnsureTextStyle(shape).BackgroundColor = ParseColor(v, default);
                        break;
                    case "TextBkgndTrans":
                        EnsureTextStyle(shape).BackgroundTransparency = ParseDouble(v);
                        break;
                    case "TxtPinX":
                        EnsureTextStyle(shape).TextPinX = ParseDouble(v);
                        break;
                    case "TxtPinY":
                        EnsureTextStyle(shape).TextPinY = ParseDouble(v);
                        break;
                    case "TxtWidth":
                        EnsureTextStyle(shape).TextWidth = ParseDouble(v);
                        break;
                    case "TxtHeight":
                        EnsureTextStyle(shape).TextHeight = ParseDouble(v);
                        break;
                    case "TxtLocPinX":
                        EnsureTextStyle(shape).TextLocPinX = ParseDouble(v);
                        break;
                    case "TxtLocPinY":
                        EnsureTextStyle(shape).TextLocPinY = ParseDouble(v);
                        break;
                    case "TxtAngle":
                        EnsureTextStyle(shape).TextAngle = ParseDouble(v);
                        break;
                    case "LayerMember":
                        ParseLayerIndexes(v, shape.LayerIndexes);
                        break;
                    case "Relationships":
                        shape.RelationshipsValue = v;
                        shape.RelationshipsFormula = cell.Attribute("F")?.Value;
                        break;
                    case "ShapePlaceStyle":
                        if (TryParseCellIntValue(v, out int shapePlacementStyle) &&
                            Enum.IsDefined(typeof(VisioPlacementStyle), shapePlacementStyle)) {
                            shape.PlacementStyle = (VisioPlacementStyle)shapePlacementStyle;
                        } else {
                            shape.PreservedCellElements.Add(new XElement(cell));
                        }
                        break;
                    case "ShapePlaceFlip":
                        if (TryParseCellIntValue(v, out int shapePlacementFlip) &&
                            IsValidPlacementFlip(shapePlacementFlip)) {
                            shape.PlacementFlip = (VisioPlacementFlip)shapePlacementFlip;
                        } else {
                            shape.PreservedCellElements.Add(new XElement(cell));
                        }
                        break;
                    case "ShapePlowCode":
                        if (TryParseCellIntValue(v, out int shapePlowCode) &&
                            Enum.IsDefined(typeof(VisioShapePlowCode), shapePlowCode)) {
                            shape.PlowCode = (VisioShapePlowCode)shapePlowCode;
                        } else {
                            shape.PreservedCellElements.Add(new XElement(cell));
                        }
                        break;
                    case "ShapePermeablePlace":
                        shape.AllowPlacementOnTop = TryParseTruthyCellValue(v);
                        break;
                    case "ShapePermeableX":
                        shape.AllowHorizontalConnectorRoutingThrough = TryParseTruthyCellValue(v);
                        break;
                    case "ShapePermeableY":
                        shape.AllowVerticalConnectorRoutingThrough = TryParseTruthyCellValue(v);
                        break;
                    case "ShapeSplit":
                        shape.CanSplitShapes = TryParseTruthyCellValue(v);
                        break;
                    case "ShapeSplittable":
                        shape.CanBeSplit = TryParseTruthyCellValue(v);
                        break;
                    default:
                        if (VisioProtection.IsCellName(n) &&
                            shape.Protection.TrySetCellValue(n, ParseNullableBoolCell(v))) {
                            break;
                        }

                        if (ShouldPreserveShapeCell(n)) {
                            shape.PreservedCellElements.Add(new XElement(cell));
                        }
                        break;
                }
            }

            XElement? xform = shapeElement.Element(ns + "XForm") ?? shapeElement.Element(ns + "XForm1D");
            if (xform != null) {
                if (!pinXFound) {
                    XElement? pinX = xform.Element(ns + "PinX");
                    if (pinX != null) {
                        shape.PinX = ParseDouble(pinX.Value);
                        pinXFound = true;
                    }
                }
                if (!pinYFound) {
                    XElement? pinY = xform.Element(ns + "PinY");
                    if (pinY != null) {
                        shape.PinY = ParseDouble(pinY.Value);
                        pinYFound = true;
                    }
                }
                if (!widthFound) {
                    XElement? width = xform.Element(ns + "Width");
                    if (width != null) {
                        shape.Width = ParseDouble(width.Value);
                        widthFound = true;
                        shape.HasExplicitWidth = true;
                    }
                }
                if (!heightFound) {
                    XElement? height = xform.Element(ns + "Height");
                    if (height != null) {
                        shape.Height = ParseDouble(height.Value);
                        heightFound = true;
                        shape.HasExplicitHeight = true;
                    }
                }
                if (!locPinXFound) {
                    XElement? locPinX = xform.Element(ns + "LocPinX");
                    if (locPinX != null) {
                        shape.LocPinX = ParseDouble(locPinX.Value);
                        locPinXFound = true;
                        shape.HasExplicitLocPinX = true;
                    }
                }
                if (!locPinYFound) {
                    XElement? locPinY = xform.Element(ns + "LocPinY");
                    if (locPinY != null) {
                        shape.LocPinY = ParseDouble(locPinY.Value);
                        locPinYFound = true;
                        shape.HasExplicitLocPinY = true;
                    }
                }
                if (!angleFound) {
                    XElement? angle = xform.Element(ns + "Angle");
                    if (angle != null) {
                        shape.Angle = ParseDouble(angle.Value);
                        angleFound = true;
                    }
                }
            }

            if (!locPinXFound) {
                shape.LocPinX = shape.Width / 2;
            }
            if (!locPinYFound) {
                shape.LocPinY = shape.Height / 2;
            }
            if (!angleFound) {
                shape.Angle = 0;
            }
            if (!lineWeightFound) {
                shape.LineWeight = DefaultLineWeight;
            }
        }

        private static void ParseShapeProperties(VisioShape shape, XElement shapeElement, XNamespace ns, IReadOnlyDictionary<int, string>? faceNamesById) {
            List<XElement> sectionElements = shapeElement.Elements(ns + "Section").ToList();

            XElement? charSection = sectionElements.FirstOrDefault(e => IsCharacterSection(e.Attribute("N")?.Value));
            if (charSection != null && TryParseSimpleCharSection(shape, charSection, ns, faceNamesById)) {
                shape.HasModeledCharSection = true;
            }

            XElement? paraSection = sectionElements.FirstOrDefault(e => IsParagraphSection(e.Attribute("N")?.Value));
            if (paraSection != null && TryParseSimpleParaSection(shape, paraSection, ns)) {
                shape.HasModeledParaSection = true;
            }

            foreach (XElement geometrySection in sectionElements.Where(section =>
                         string.Equals(section.Attribute("N")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))) {
                shape.PreservedGeometrySections.Add(new XElement(geometrySection));
            }
            foreach (XElement section in sectionElements.Where(section => ShouldPreserveShapeSection(shape, section))) {
                shape.PreservedNonGeometrySections.Add(new XElement(section));
            }

            XElement? hyperlinkSection = sectionElements.FirstOrDefault(e => e.Attribute("N")?.Value == "Hyperlink");
            if (hyperlinkSection != null) {
                ParseHyperlinks(hyperlinkSection, ns, shape.Hyperlinks);
            }

            XElement? userSection = sectionElements.FirstOrDefault(e => e.Attribute("N")?.Value == "User");
            if (userSection != null) {
                ParseUserCells(userSection, ns, shape.UserCells);
            }

            XElement? connectionSection = sectionElements.FirstOrDefault(e => e.Attribute("N")?.Value == "Connection");
            if (connectionSection != null) {
                foreach (XElement row in connectionSection.Elements(ns + "Row")) {
                    double x = 0;
                    double y = 0;
                    double dirX = 0;
                    double dirY = 0;
                    int? sectionIndex = null;
                    if (int.TryParse(row.Attribute("IX")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedSectionIndex) &&
                        parsedSectionIndex >= 0) {
                        sectionIndex = parsedSectionIndex;
                    }

                    foreach (XElement cell in row.Elements(ns + "Cell")) {
                        string? n = cell.Attribute("N")?.Value;
                        string? v = cell.Attribute("V")?.Value;
                        switch (n) {
                            case "X":
                                x = ParseDouble(v);
                                break;
                            case "Y":
                                y = ParseDouble(v);
                                break;
                            case "DirX":
                                dirX = ParseDouble(v);
                                break;
                            case "DirY":
                                dirY = ParseDouble(v);
                                break;
                        }
                    }
                    shape.ConnectionPoints.Add(new VisioConnectionPoint(x, y, dirX, dirY) {
                        SectionIndex = sectionIndex
                    });
                }
            }

            XElement? propSection = sectionElements.FirstOrDefault(e => e.Attribute("N")?.Value == "Prop");
            if (propSection != null) {
                ParseShapeDataRows(propSection, ns, shape);
            }
        }

        private static string? GetOriginalId(XElement shapeElement, XNamespace ns) {
            XElement? propSection = shapeElement.Elements(ns + "Section")
                .FirstOrDefault(e => e.Attribute("N")?.Value == "Prop");
            if (propSection == null) {
                return null;
            }

            XElement? originalIdRow = propSection.Elements(ns + "Row")
                .FirstOrDefault(row => string.Equals(row.Attribute("N")?.Value, OriginalIdPropName, StringComparison.Ordinal));
            return originalIdRow?.Elements(ns + "Cell")
                .FirstOrDefault(c => c.Attribute("N")?.Value == "Value")
                ?.Attribute("V")?.Value;
        }

        private static void ParseChildShapes(VisioShape shape, XElement shapeElement, XNamespace ns, IReadOnlyDictionary<int, string>? faceNamesById, int depth) {
            XElement? childShapes = shapeElement.Element(ns + "Shapes");
            if (childShapes == null) {
                return;
            }

            List<XElement> childElements = childShapes.Elements(ns + "Shape").ToList();
            foreach (XElement childElement in childElements) {
                VisioShape childShape = ParseShapeCore(childElement, ns, faceNamesById, shape, depth + 1);
                shape.Children.Add(childShape);
            }
        }
    }
}
