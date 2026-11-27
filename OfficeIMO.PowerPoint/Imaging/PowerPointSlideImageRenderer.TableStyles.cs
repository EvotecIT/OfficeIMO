using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private enum TableCellBorderEdge {
            Left,
            Top,
            Right,
            Bottom
        }

        /// <summary>
        /// The table's style (<c>a:tblStyle</c>): the presentation's own entry with the table's style id, else the built-in
        /// entry; none when the table names no style or neither has it.
        /// </summary>
        internal static OfficeXmlNode ResolveTableStyleElement(PowerPointTable table) {
            string? styleId = table.StyleId;
            if (string.IsNullOrWhiteSpace(styleId)) {
                return default;
            }

            string resolvedStyleId = styleId!;
            PowerPointPartRef presentationPart = ResolveTablePresentationPart(table);
            if (!presentationPart.IsNone) {
                OfficeXmlNode packageStyle = FindTableStyle(
                    presentationPart.RelatedPart(PowerPointDocument.TableStylesRelationshipType).Root, resolvedStyleId);
                if (!packageStyle.IsNone) {
                    return packageStyle;
                }
            }

            // The built-in styles are parsed for each lookup, so callers never share or edit one tree.
            byte[] builtIn = PowerPointUtils.TableStylesBytes.Value;
            OfficeXmlTree tree;
            try {
                tree = OfficeXmlTree.Parse(builtIn, builtIn.Length, new OfficeXmlNamespaces.Table());
            } catch (System.Xml.XmlException) {
                return default;
            }
            return FindTableStyle(new OfficeXmlNode(tree, tree.Root), resolvedStyleId);
        }

        /// <summary>The presentation part of the table's slide, or of the part that holds a table of a layout or master.</summary>
        private static PowerPointPartRef ResolveTablePresentationPart(PowerPointTable table) {
            PowerPointDocument? document = table.OwnerSlide?.OwnerDocument ?? table.OwnerPartRef.Document;
            return document == null ? default : document.PresentationPartRef;
        }

        private static OfficeXmlNode FindTableStyle(OfficeXmlNode styleList, string styleId) {
            foreach (OfficeXmlNode style in styleList.Children(D, "tblStyle")) {
                if (string.Equals(style.Attribute("styleId"), styleId, StringComparison.OrdinalIgnoreCase)) {
                    return style;
                }
            }

            return default;
        }

        private static OfficeColor? ResolveTableStyleFillColor(PowerPointTable table, int row, int column, OfficeXmlNode tableStyle, OfficeXmlNode colorScheme) {
            foreach (OfficeXmlNode region in EnumerateApplicableTableStyleRegions(table, row, column, tableStyle)) {
                OfficeXmlNode fill = region.Child(D, "tcStyle").Child(D, "fill");
                if (!fill.Child(D, "noFill").IsNone) return OfficeColor.Transparent;
                OfficeXmlNode solidFill = fill.Child(D, "solidFill");
                OfficeColor? color = OfficeXmlThemeColors.ResolveColor(solidFill, colorScheme);
                if (color.HasValue) {
                    return color;
                }
            }

            return null;
        }

        /// <summary>The <c>a:tcTxStyle</c> of each style region that applies to the cell, nearest first.</summary>
        internal static List<OfficeXmlNode> ResolveTableCellTextStyleElements(
            PowerPointTable table,
            int row,
            int column,
            OfficeXmlNode tableStyle) {
            var styles = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode region in EnumerateApplicableTableStyleRegions(table, row, column, tableStyle)) {
                OfficeXmlNode style = region.Child(D, "tcTxStyle");
                if (!style.IsNone) styles.Add(style);
            }
            return styles;
        }

        private static OfficeBorderSide? ResolveTableStyleBorderSide(PowerPointTable table, int row, int column, int rowSpan, int columnSpan, OfficeXmlNode tableStyle, TableCellBorderEdge edge, OfficeXmlNode colorScheme) {
            foreach (OfficeXmlNode region in EnumerateApplicableTableStyleRegions(table, row, column, tableStyle)) {
                OfficeXmlNode borders = region.Child(D, "tcStyle").Child(D, "tcBdr");
                OfficeXmlNode outline = ResolveTableStyleBorderOutline(table, row, column, rowSpan, columnSpan, borders, edge);
                if (!outline.IsNone) {
                    return ResolveTableCellBorderSide(outline, colorScheme);
                }
            }

            return null;
        }

        private static OfficeXmlNode ResolveTableStyleBorderOutline(PowerPointTable table, int row, int column, int rowSpan, int columnSpan, OfficeXmlNode borders, TableCellBorderEdge edge) {
            if (borders.IsNone) {
                return default;
            }

            bool firstRow = row == 0;
            bool lastRow = row + Math.Max(1, rowSpan) >= table.RowCount;
            bool firstColumn = column == 0;
            bool lastColumn = column + Math.Max(1, columnSpan) >= table.ColumnCount;
            return edge switch {
                TableCellBorderEdge.Left => firstColumn
                    ? BorderOutline(borders, "left")
                    : FirstOutline(BorderOutline(borders, "insideV"), BorderOutline(borders, "left")),
                TableCellBorderEdge.Top => firstRow
                    ? BorderOutline(borders, "top")
                    : FirstOutline(BorderOutline(borders, "insideH"), BorderOutline(borders, "top")),
                TableCellBorderEdge.Right => lastColumn
                    ? BorderOutline(borders, "right")
                    : FirstOutline(BorderOutline(borders, "insideV"), BorderOutline(borders, "right")),
                TableCellBorderEdge.Bottom => lastRow
                    ? BorderOutline(borders, "bottom")
                    : FirstOutline(BorderOutline(borders, "insideH"), BorderOutline(borders, "bottom")),
                _ => default
            };
        }

        /// <summary>The <c>a:ln</c> of one side of a table style's <c>a:tcBdr</c>, or none.</summary>
        private static OfficeXmlNode BorderOutline(OfficeXmlNode borders, string side) =>
            borders.Child(D, side).Child(D, "ln");

        private static OfficeXmlNode FirstOutline(OfficeXmlNode preferred, OfficeXmlNode fallback) =>
            preferred.IsNone ? fallback : preferred;

        private static IEnumerable<OfficeXmlNode> EnumerateApplicableTableStyleRegions(PowerPointTable table, int row, int column, OfficeXmlNode tableStyle) {
            if (tableStyle.IsNone) {
                yield break;
            }

            bool firstRow = row == 0;
            bool lastRow = row == table.RowCount - 1;
            bool firstColumn = column == 0;
            bool lastColumn = column == table.ColumnCount - 1;
            OfficeXmlNode region;

            if (table.FirstRow && table.FirstColumn && firstRow && firstColumn && !(region = tableStyle.Child(D, "nwCell")).IsNone) {
                yield return region;
            }

            if (table.FirstRow && table.LastColumn && firstRow && lastColumn && !(region = tableStyle.Child(D, "neCell")).IsNone) {
                yield return region;
            }

            if (table.LastRow && table.FirstColumn && lastRow && firstColumn && !(region = tableStyle.Child(D, "swCell")).IsNone) {
                yield return region;
            }

            if (table.LastRow && table.LastColumn && lastRow && lastColumn && !(region = tableStyle.Child(D, "seCell")).IsNone) {
                yield return region;
            }

            if (table.FirstRow && firstRow && !(region = tableStyle.Child(D, "firstRow")).IsNone) {
                yield return region;
            }

            if (table.LastRow && lastRow && !(region = tableStyle.Child(D, "lastRow")).IsNone) {
                yield return region;
            }

            if (table.FirstColumn && firstColumn && !(region = tableStyle.Child(D, "firstCol")).IsNone) {
                yield return region;
            }

            if (table.LastColumn && lastColumn && !(region = tableStyle.Child(D, "lastCol")).IsNone) {
                yield return region;
            }

            if (table.BandedRows && !IsExcludedTableStyleBandRow(table, row)) {
                OfficeXmlNode band = ResolveHorizontalBand(table, row, tableStyle);
                if (!band.IsNone) {
                    yield return band;
                }
            }

            if (table.BandedColumns && !IsExcludedTableStyleBandColumn(table, column)) {
                OfficeXmlNode band = ResolveVerticalBand(table, column, tableStyle);
                if (!band.IsNone) {
                    yield return band;
                }
            }

            if (!(region = tableStyle.Child(D, "wholeTbl")).IsNone) {
                yield return region;
            }
        }

        private static bool IsExcludedTableStyleBandRow(PowerPointTable table, int row) =>
            (table.FirstRow && row == 0) || (table.LastRow && row == table.RowCount - 1);

        private static bool IsExcludedTableStyleBandColumn(PowerPointTable table, int column) =>
            (table.FirstColumn && column == 0) || (table.LastColumn && column == table.ColumnCount - 1);

        private static OfficeXmlNode ResolveHorizontalBand(PowerPointTable table, int row, OfficeXmlNode tableStyle) {
            int visualRow = row - (table.FirstRow ? 1 : 0);
            return visualRow % 2 == 0
                ? tableStyle.Child(D, "band1H")
                : tableStyle.Child(D, "band2H");
        }

        private static OfficeXmlNode ResolveVerticalBand(PowerPointTable table, int column, OfficeXmlNode tableStyle) {
            int visualColumn = column - (table.FirstColumn ? 1 : 0);
            return visualColumn % 2 == 0
                ? tableStyle.Child(D, "band1V")
                : tableStyle.Child(D, "band2V");
        }
    }
}
