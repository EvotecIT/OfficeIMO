using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Writes the formatting tables described by BIFF12, retaining cell-format and font indices.</summary>
    internal static class XlsbStylesheetMarkup {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;

        private static readonly string[] FillPatterns = {
            "none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown",
            "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp",
            "lightGrid", "lightTrellis", "gray125", "gray0625"
        };
        private static readonly string?[] BorderStyles = {
            null, "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed",
            "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"
        };
        private static readonly string?[] HorizontalAlignments = {
            null, "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"
        };
        private static readonly string?[] VerticalAlignments = { "top", "center", null, "justify", "distributed" };
        private static readonly string[] ApplyAttributes = {
            "applyNumberFormat", "applyFont", "applyAlignment", "applyBorder", "applyFill", "applyProtection"
        };

        internal static void Write(OfficeXmlNode sheet, XlsbStylesheet source) {
            if (source.NumberFormats.Count != 0) {
                OfficeXmlNode formats = AddCounted(sheet, "numFmts", source.NumberFormats.Count);
                foreach (var format in source.NumberFormats.OrderBy(item => item.Key)) {
                    OfficeXmlNode node = Add(formats, "numFmt");
                    node.SetUInt32Attribute("numFmtId", format.Key);
                    node.SetAttribute("formatCode", format.Value);
                }
            }

            OfficeXmlNode fonts = AddCounted(sheet, "fonts", source.Fonts.Count);
            foreach (XlsbFont font in source.Fonts) WriteFont(Add(fonts, "font"), font, run: false);

            OfficeXmlNode fills = AddCounted(sheet, "fills", source.Fills.Count);
            foreach (XlsbFill fill in source.Fills) {
                // Gradient stops stay in the native part. The projection represents their base colors.
                OfficeXmlNode pattern = Add(Add(fills, "fill"), "patternFill");
                pattern.SetAttribute("patternType", FillPattern(fill.Pattern));
                if (fill.Foreground != null) WriteColor(Add(pattern, "fgColor"), fill.Foreground);
                if (fill.Background != null) WriteColor(Add(pattern, "bgColor"), fill.Background);
            }

            OfficeXmlNode borders = AddCounted(sheet, "borders", source.Borders.Count);
            foreach (XlsbBorder border in source.Borders) WriteBorder(Add(borders, "border"), border);

            OfficeXmlNode styleFormats = AddCounted(sheet, "cellStyleXfs", source.CellStyleFormats.Count);
            foreach (XlsbCellFormat format in source.CellStyleFormats) WriteCellFormat(Add(styleFormats, "xf"), format, styleFormat: true);
            OfficeXmlNode cellFormats = AddCounted(sheet, "cellXfs", source.CellFormats.Count);
            foreach (XlsbCellFormat format in source.CellFormats) WriteCellFormat(Add(cellFormats, "xf"), format, styleFormat: false);
            WriteCellStyles(sheet, source.CellStyles);

            OfficeXmlNode differentialFormats = AddCounted(sheet, "dxfs", source.DifferentialFormatCount);
            foreach (OfficeXmlNode differential in source.ReadDifferentialFormatMarkup()) differentialFormats.Append(differential);
            OfficeXmlNode tableStyles = AddCounted(sheet, "tableStyles", 0);
            tableStyles.SetAttribute("defaultTableStyle", "TableStyleMedium2");
            tableStyles.SetAttribute("defaultPivotStyle", "PivotStyleLight16");
        }

        internal static string FillPattern(uint value) => value < FillPatterns.Length ? FillPatterns[value] : "none";

        internal static string? BorderStyle(byte value) => value < BorderStyles.Length ? BorderStyles[value] : null;

        /// <summary>Writes a font record as either a cell font or the font properties of an inline text run.</summary>
        internal static void WriteFont(OfficeXmlNode node, XlsbFont source, bool run) {
            if (source.Weight >= 700) Add(node, "b");
            if ((source.Flags & 0x0002) != 0) Add(node, "i");
            if ((source.Flags & 0x0008) != 0) Add(node, "strike");
            if ((source.Flags & 0x0010) != 0) Add(node, "outline");
            if ((source.Flags & 0x0020) != 0) Add(node, "shadow");
            if ((source.Flags & 0x0040) != 0) Add(node, "condense");
            if ((source.Flags & 0x0080) != 0) Add(node, "extend");
            string? underline = source.Underline switch {
                1 => "single", 2 => "double", 0x21 => "singleAccounting", 0x22 => "doubleAccounting", _ => null
            };
            if (underline != null) SetValue(node, "u", underline);
            if (source.Script is 1 or 2) SetValue(node, "vertAlign", source.Script == 1 ? "superscript" : "subscript");
            Add(node, "sz").SetDoubleAttribute("val", source.HeightTwips / 20D);
            if (source.Color != null) WriteColor(Add(node, "color"), source.Color);
            SetValue(node, run ? "rFont" : "name", source.Name);
            if (source.Family != 0) Add(node, "family").SetUInt32Attribute("val", source.Family);
            if (source.CharacterSet != 1) Add(node, "charset").SetUInt32Attribute("val", source.CharacterSet);
            if (source.Scheme is 1 or 2) SetValue(node, "scheme", source.Scheme == 1 ? "major" : "minor");
        }

        private static void WriteBorder(OfficeXmlNode node, XlsbBorder source) {
            if (source.DiagonalDown) node.SetBooleanAttribute("diagonalDown", true);
            if (source.DiagonalUp) node.SetBooleanAttribute("diagonalUp", true);
            WriteBorderSide(Add(node, "left"), source.Left);
            WriteBorderSide(Add(node, "right"), source.Right);
            WriteBorderSide(Add(node, "top"), source.Top);
            WriteBorderSide(Add(node, "bottom"), source.Bottom);
            WriteBorderSide(Add(node, "diagonal"), source.Diagonal);
        }

        private static void WriteBorderSide(OfficeXmlNode node, XlsbBorderSide source) {
            if (BorderStyle(source.Style) is string style) node.SetAttribute("style", style);
            if (source.Color != null) WriteColor(Add(node, "color"), source.Color);
        }

        private static void WriteCellFormat(OfficeXmlNode node, XlsbCellFormat source, bool styleFormat) {
            node.SetUInt32Attribute("numFmtId", source.NumberFormatId);
            node.SetUInt32Attribute("fontId", source.FontId);
            node.SetUInt32Attribute("fillId", source.FillId);
            node.SetUInt32Attribute("borderId", source.BorderId);
            for (int bit = 0; bit < ApplyAttributes.Length; bit++) {
                if ((source.ApplyFlags & (1 << bit)) != 0) node.SetBooleanAttribute(ApplyAttributes[bit], true);
            }
            if (source.PivotButton) node.SetBooleanAttribute("pivotButton", true);
            if (source.QuotePrefix) node.SetBooleanAttribute("quotePrefix", true);
            if (!styleFormat) node.SetUInt32Attribute("xfId", source.ParentFormatId);

            const byte bottomAlignment = 2;
            if (source.HorizontalAlignment != 0 || source.VerticalAlignment != bottomAlignment || source.TextRotation != 0
                || source.Indent != 0 || source.WrapText || source.JustifyLastLine || source.ShrinkToFit || source.Merged
                || source.ReadingOrder != 0 || (source.ApplyFlags & 4) != 0) {
                OfficeXmlNode alignment = Add(node, "alignment");
                if (source.HorizontalAlignment < HorizontalAlignments.Length && HorizontalAlignments[source.HorizontalAlignment] is string horizontal)
                    alignment.SetAttribute("horizontal", horizontal);
                if (source.VerticalAlignment < VerticalAlignments.Length && VerticalAlignments[source.VerticalAlignment] is string vertical)
                    alignment.SetAttribute("vertical", vertical);
                if (source.TextRotation != 0) alignment.SetUInt32Attribute("textRotation", source.TextRotation);
                if (source.Indent != 0) alignment.SetUInt32Attribute("indent", source.Indent);
                if (source.WrapText) alignment.SetBooleanAttribute("wrapText", true);
                if (source.JustifyLastLine) alignment.SetBooleanAttribute("justifyLastLine", true);
                if (source.ShrinkToFit) alignment.SetBooleanAttribute("shrinkToFit", true);
                alignment.SetAttribute("mergeCell", source.Merged ? "1" : string.Empty);
                if (source.ReadingOrder != 0) alignment.SetUInt32Attribute("readingOrder", source.ReadingOrder);
            }

            if (source.Locked || source.Hidden || (source.ApplyFlags & 32) != 0) {
                OfficeXmlNode protection = Add(node, "protection");
                protection.SetBooleanAttribute("locked", source.Locked);
                if (source.Hidden) protection.SetBooleanAttribute("hidden", true);
            }
        }

        private static void WriteCellStyles(OfficeXmlNode sheet, IReadOnlyList<XlsbCellStyle> source) {
            OfficeXmlNode styles = AddCounted(sheet, "cellStyles", Math.Max(1, source.Count));
            if (source.Count == 0) {
                OfficeXmlNode normal = Add(styles, "cellStyle");
                normal.SetAttribute("name", "Normal");
                normal.SetUInt32Attribute("xfId", 0U);
                normal.SetUInt32Attribute("builtinId", 0U);
                return;
            }
            foreach (XlsbCellStyle item in source) {
                OfficeXmlNode style = Add(styles, "cellStyle");
                style.SetAttribute("name", item.Name);
                style.SetUInt32Attribute("xfId", item.FormatId);
                if (item.BuiltIn) {
                    style.SetUInt32Attribute("builtinId", item.BuiltInId);
                    if (item.BuiltInId is 1 or 2) style.SetUInt32Attribute("iLevel", item.Level);
                }
                if (item.Hidden) style.SetBooleanAttribute("hidden", true);
                if (item.CustomBuiltIn) style.SetBooleanAttribute("customBuiltin", true);
            }
        }

        private static void WriteColor(OfficeXmlNode color, XlsbColor source) {
            switch (source.Type) {
                case 0: color.SetBooleanAttribute("auto", true); break;
                case 1: color.SetUInt32Attribute("indexed", source.Index); break;
                case 2: color.SetAttribute("rgb", $"{source.Alpha:X2}{source.Red:X2}{source.Green:X2}{source.Blue:X2}"); break;
                case 3: color.SetUInt32Attribute("theme", source.Index); break;
            }
            if (source.Tint != 0) color.SetDoubleAttribute("tint", source.Tint / (source.Tint < 0 ? 32768D : 32767D));
        }

        private static OfficeXmlNode Add(OfficeXmlNode node, string localName) => node.Append(node.CreateElement(Spreadsheet, localName));

        private static OfficeXmlNode AddCounted(OfficeXmlNode parent, string localName, int count) {
            OfficeXmlNode node = Add(parent, localName);
            node.SetUInt32Attribute("count", checked((uint)count));
            return node;
        }

        private static void SetValue(OfficeXmlNode parent, string localName, string value) => Add(parent, localName).SetAttribute("val", value);
    }
}
