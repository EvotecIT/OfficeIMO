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

        private static VisioTextStyle EnsureTextStyle(VisioShape shape) {
            if (shape.TextStyle == null) {
                shape.TextStyle = new VisioTextStyle();
            }

            return shape.TextStyle;
        }

        private static bool TryParseSimpleCharSection(VisioShape shape, XElement section, XNamespace ns, IReadOnlyDictionary<int, string>? faceNamesById) {
            List<XElement> rows = section.Elements(ns + "Row").ToList();
            if (rows.Count != 1) {
                return false;
            }

            int? fontFaceId = null;
            string? fontFamily = null;
            OfficeIMO.Drawing.OfficeColor? color = null;
            double? size = null;
            bool? bold = null;
            bool? italic = null;
            bool? underline = null;
            foreach (XElement cell in rows[0].Elements(ns + "Cell")) {
                string? name = cell.Attribute("N")?.Value;
                string? value = cell.Attribute("V")?.Value;
                switch (name) {
                    case "Font":
                        if (!TryParseCellIntValue(value, out int parsedFontFaceId)) {
                            return false;
                        }

                        if (faceNamesById == null || !faceNamesById.TryGetValue(parsedFontFaceId, out string? resolvedFontFamily)) {
                            return false;
                        }

                        fontFaceId = parsedFontFaceId;
                        fontFamily = resolvedFontFamily;
                        break;
                    case "Color":
                        color = ParseColor(value, default);
                        break;
                    case "Size":
                        size = ParseTextSizeCell(cell);
                        break;
                    case "Style":
                        if (!TryParseCellIntValue(value, out int styleValue)) {
                            return false;
                        }

                        bold = (styleValue & 1) != 0;
                        italic = (styleValue & 2) != 0;
                        underline = (styleValue & 4) != 0;
                        break;
                    default:
                        return false;
                }
            }

            VisioTextStyle textStyle = EnsureTextStyle(shape);
            textStyle.FontFaceId = fontFaceId;
            textStyle.FontFamily = fontFamily;
            textStyle.Color = color;
            textStyle.Size = size;
            textStyle.Bold = bold;
            textStyle.Italic = italic;
            textStyle.Underline = underline;
            return true;
        }

        private static bool TryParseSimpleParaSection(VisioShape shape, XElement section, XNamespace ns) {
            List<XElement> rows = section.Elements(ns + "Row").ToList();
            if (rows.Count != 1) {
                return false;
            }

            VisioTextHorizontalAlignment? horizontalAlignment = null;
            foreach (XElement cell in rows[0].Elements(ns + "Cell")) {
                string? name = cell.Attribute("N")?.Value;
                string? value = cell.Attribute("V")?.Value;
                switch (name) {
                    case "HorzAlign":
                        if (TryParseCellIntValue(value, out int horizontalAlign) &&
                            Enum.IsDefined(typeof(VisioTextHorizontalAlignment), horizontalAlign)) {
                            horizontalAlignment = (VisioTextHorizontalAlignment)horizontalAlign;
                        } else {
                            return false;
                        }

                        break;
                    default:
                        return false;
                }
            }

            EnsureTextStyle(shape).HorizontalAlignment = horizontalAlignment;
            return true;
        }

        private static bool TryParseSimpleConnectorCharSection(VisioConnector connector, XElement section, XNamespace ns, IReadOnlyDictionary<int, string>? faceNamesById) {
            List<XElement> rows = section.Elements(ns + "Row").ToList();
            if (rows.Count != 1) {
                return false;
            }

            int? fontFaceId = null;
            string? fontFamily = null;
            OfficeIMO.Drawing.OfficeColor? color = null;
            double? size = null;
            bool? bold = null;
            bool? italic = null;
            bool? underline = null;
            foreach (XElement cell in rows[0].Elements(ns + "Cell")) {
                string? name = cell.Attribute("N")?.Value;
                string? value = cell.Attribute("V")?.Value;
                switch (name) {
                    case "Font":
                        if (!TryParseCellIntValue(value, out int parsedFontFaceId)) {
                            return false;
                        }

                        if (faceNamesById == null || !faceNamesById.TryGetValue(parsedFontFaceId, out string? resolvedFontFamily)) {
                            return false;
                        }

                        fontFaceId = parsedFontFaceId;
                        fontFamily = resolvedFontFamily;
                        break;
                    case "Color":
                        color = ParseColor(value, default);
                        break;
                    case "Size":
                        size = ParseTextSizeCell(cell);
                        break;
                    case "Style":
                        if (!TryParseCellIntValue(value, out int styleValue)) {
                            return false;
                        }

                        bold = (styleValue & 1) != 0;
                        italic = (styleValue & 2) != 0;
                        underline = (styleValue & 4) != 0;
                        break;
                    default:
                        return false;
                }
            }

            VisioTextStyle textStyle = EnsureConnectorTextStyle(connector);
            textStyle.FontFaceId = fontFaceId;
            textStyle.FontFamily = fontFamily;
            textStyle.Color = color;
            textStyle.Size = size;
            textStyle.Bold = bold;
            textStyle.Italic = italic;
            textStyle.Underline = underline;
            return true;
        }

        private static double ParseTextSizeCell(XElement cell) {
            double size = ParseDouble(cell.Attribute("V")?.Value);
            string? unit = cell.Attribute("U")?.Value;
            if (string.Equals(unit, "PT", StringComparison.OrdinalIgnoreCase) && size <= 3D) {
                return Math.Round(size * 72D, 10);
            }

            return size;
        }

        private static bool TryParseSimpleConnectorParaSection(VisioConnector connector, XElement section, XNamespace ns) {
            List<XElement> rows = section.Elements(ns + "Row").ToList();
            if (rows.Count != 1) {
                return false;
            }

            VisioTextHorizontalAlignment? horizontalAlignment = null;
            foreach (XElement cell in rows[0].Elements(ns + "Cell")) {
                string? name = cell.Attribute("N")?.Value;
                string? value = cell.Attribute("V")?.Value;
                switch (name) {
                    case "HorzAlign":
                        if (TryParseCellIntValue(value, out int horizontalAlign) &&
                            Enum.IsDefined(typeof(VisioTextHorizontalAlignment), horizontalAlign)) {
                            horizontalAlignment = (VisioTextHorizontalAlignment)horizontalAlign;
                        } else {
                            return false;
                        }

                        break;
                    default:
                        return false;
                }
            }

            EnsureConnectorTextStyle(connector).HorizontalAlignment = horizontalAlignment;
            return true;
        }
    }
}
