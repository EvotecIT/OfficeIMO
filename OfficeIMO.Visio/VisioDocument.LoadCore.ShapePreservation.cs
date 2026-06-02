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

        private static void CaptureShapeChildOrder(VisioShape shape, XElement shapeElement) {
            shape.PreservedShapeChildren.Clear();
            foreach (XElement child in shapeElement.Elements()) {
                string localName = child.Name.LocalName;
                if (string.Equals(localName, "XForm", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(localName, "XForm1D", StringComparison.OrdinalIgnoreCase)) {
                    shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("XForm"));
                    continue;
                }

                if (string.Equals(localName, "Cell", StringComparison.OrdinalIgnoreCase)) {
                    string? cellName = child.Attribute("N")?.Value;
                    if (IsModeledShapeCell(cellName)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry($"Cell:{cellName}"));
                    } else {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry(child));
                    }

                    continue;
                }

                if (string.Equals(localName, "Section", StringComparison.OrdinalIgnoreCase)) {
                    string? sectionName = child.Attribute("N")?.Value;
                    if (string.Equals(sectionName, "Geometry", StringComparison.OrdinalIgnoreCase)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:Geometry"));
                    } else if (string.Equals(sectionName, "Connection", StringComparison.OrdinalIgnoreCase)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:Connection"));
                    } else if (string.Equals(sectionName, "User", StringComparison.OrdinalIgnoreCase)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:User"));
                    } else if (shape.HasModeledCharSection && IsCharacterSection(sectionName)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:Char"));
                    } else if (shape.HasModeledParaSection && IsParagraphSection(sectionName)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:Para"));
                    } else if (string.Equals(sectionName, "Hyperlink", StringComparison.OrdinalIgnoreCase)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:Hyperlink"));
                    } else if (string.Equals(sectionName, "Prop", StringComparison.OrdinalIgnoreCase)) {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Section:Prop"));
                    } else {
                        shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry(child));
                    }

                    continue;
                }

                if (string.Equals(localName, "Text", StringComparison.OrdinalIgnoreCase)) {
                    shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Text"));
                    continue;
                }

                if (string.Equals(localName, "Shapes", StringComparison.OrdinalIgnoreCase)) {
                    shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry("Shapes"));
                    continue;
                }

                shape.PreservedShapeChildren.Add(new VisioShape.PreservedShapeChildEntry(child));
            }
        }

        private static bool IsModeledShapeCell(string? cellName) {
            return string.Equals(cellName, "PinX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "PinY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "Width", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "Height", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LocPinX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LocPinY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "Angle", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LineWeight", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LinePattern", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LineColor", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "FillPattern", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "FillForegnd", StringComparison.OrdinalIgnoreCase) ||
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
                   string.Equals(cellName, "TxtAngle", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "LayerMember", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "Relationships", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ObjType", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapePlaceStyle", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapePlaceFlip", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapePlowCode", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapePermeablePlace", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapePermeableX", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapePermeableY", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapeSplit", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(cellName, "ShapeSplittable", StringComparison.OrdinalIgnoreCase) ||
                   VisioProtection.IsCellName(cellName);
        }

        private static bool ShouldPreserveShapeCell(string? cellName) {
            return !string.IsNullOrWhiteSpace(cellName) &&
                   !string.Equals(cellName, "PinX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PinY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "Width", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "Height", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LocPinX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LocPinY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "Angle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineWeight", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LinePattern", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "FillPattern", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineColor", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "FillForegnd", StringComparison.OrdinalIgnoreCase) &&
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
                   !string.Equals(cellName, "TxtAngle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LayerMember", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "Relationships", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ObjType", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapePlaceStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapePlaceFlip", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapePlowCode", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapePermeablePlace", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapePermeableX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapePermeableY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapeSplit", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapeSplittable", StringComparison.OrdinalIgnoreCase) &&
                   !VisioProtection.IsCellName(cellName);
        }

        private static bool ShouldPreserveShapeSection(XElement section) {
            string? sectionName = section.Attribute("N")?.Value;
            return !string.Equals(sectionName, "Geometry", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(sectionName, "Connection", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(sectionName, "User", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(sectionName, "Hyperlink", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(sectionName, "Prop", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveShapeSection(VisioShape shape, XElement section) {
            string? sectionName = section.Attribute("N")?.Value;
            if (shape.HasModeledCharSection && IsCharacterSection(sectionName)) {
                return false;
            }

            if (shape.HasModeledParaSection && IsParagraphSection(sectionName)) {
                return false;
            }

            return ShouldPreserveShapeSection(section);
        }

        private static bool IsCharacterSection(string? sectionName) {
            return string.Equals(sectionName, "Character", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(sectionName, "Char", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsParagraphSection(string? sectionName) {
            return string.Equals(sectionName, "Paragraph", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(sectionName, "Para", StringComparison.OrdinalIgnoreCase);
        }
    }
}
