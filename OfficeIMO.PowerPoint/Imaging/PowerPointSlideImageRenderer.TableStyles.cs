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

            OfficeXmlTree? tree = BuiltInTableStyles.Value;
            return tree == null ? default : FindTableStyle(new OfficeXmlNode(tree, tree.Root), resolvedStyleId);
        }

        // The built-in styles, parsed once. The tree is shared, so what comes from it is only ever read, never edited.
        private static readonly Lazy<OfficeXmlTree?> BuiltInTableStyles = new(() => {
            byte[] builtIn = PowerPointUtils.TableStylesBytes.Value;
            try {
                return OfficeXmlTree.Parse(builtIn, builtIn.Length, new OfficeXmlNamespaces.Table());
            } catch (System.Xml.XmlException) {
                return null;
            }
        });

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

            TableStyleFlags flags = StyleFlags(table);
            bool firstRow = row == 0;
            bool lastRow = row == flags.RowCount - 1;
            bool firstColumn = column == 0;
            bool lastColumn = column == flags.ColumnCount - 1;
            OfficeXmlNode region;

            if (flags.FirstRow && flags.FirstColumn && firstRow && firstColumn && !(region = tableStyle.Child(D, "nwCell")).IsNone) {
                yield return region;
            }

            if (flags.FirstRow && flags.LastColumn && firstRow && lastColumn && !(region = tableStyle.Child(D, "neCell")).IsNone) {
                yield return region;
            }

            if (flags.LastRow && flags.FirstColumn && lastRow && firstColumn && !(region = tableStyle.Child(D, "swCell")).IsNone) {
                yield return region;
            }

            if (flags.LastRow && flags.LastColumn && lastRow && lastColumn && !(region = tableStyle.Child(D, "seCell")).IsNone) {
                yield return region;
            }

            if (flags.FirstRow && firstRow && !(region = tableStyle.Child(D, "firstRow")).IsNone) {
                yield return region;
            }

            if (flags.LastRow && lastRow && !(region = tableStyle.Child(D, "lastRow")).IsNone) {
                yield return region;
            }

            if (flags.FirstColumn && firstColumn && !(region = tableStyle.Child(D, "firstCol")).IsNone) {
                yield return region;
            }

            if (flags.LastColumn && lastColumn && !(region = tableStyle.Child(D, "lastCol")).IsNone) {
                yield return region;
            }

            if (flags.BandedRows && !IsExcludedTableStyleBandRow(flags, row)) {
                OfficeXmlNode band = ResolveHorizontalBand(flags, row, tableStyle);
                if (!band.IsNone) {
                    yield return band;
                }
            }

            if (flags.BandedColumns && !IsExcludedTableStyleBandColumn(flags, column)) {
                OfficeXmlNode band = ResolveVerticalBand(flags, column, tableStyle);
                if (!band.IsNone) {
                    yield return band;
                }
            }

            if (!(region = tableStyle.Child(D, "wholeTbl")).IsNone) {
                yield return region;
            }
        }

        private static bool IsExcludedTableStyleBandRow(TableStyleFlags flags, int row) =>
            (flags.FirstRow && row == 0) || (flags.LastRow && row == flags.RowCount - 1);

        private static bool IsExcludedTableStyleBandColumn(TableStyleFlags flags, int column) =>
            (flags.FirstColumn && column == 0) || (flags.LastColumn && column == flags.ColumnCount - 1);

        private static OfficeXmlNode ResolveHorizontalBand(TableStyleFlags flags, int row, OfficeXmlNode tableStyle) {
            int visualRow = row - (flags.FirstRow ? 1 : 0);
            return visualRow % 2 == 0
                ? tableStyle.Child(D, "band1H")
                : tableStyle.Child(D, "band2H");
        }

        private static OfficeXmlNode ResolveVerticalBand(TableStyleFlags flags, int column, OfficeXmlNode tableStyle) {
            int visualColumn = column - (flags.FirstColumn ? 1 : 0);
            return visualColumn % 2 == 0
                ? tableStyle.Child(D, "band1V")
                : tableStyle.Child(D, "band2V");
        }

        /// <summary>The style options of a table and its size, which style regions are chosen by.</summary>
        private readonly struct TableStyleFlags {
            internal TableStyleFlags(PowerPointTable table) {
                FirstRow = table.FirstRow;
                LastRow = table.LastRow;
                FirstColumn = table.FirstColumn;
                LastColumn = table.LastColumn;
                BandedRows = table.BandedRows;
                BandedColumns = table.BandedColumns;
                RowCount = table.RowCount;
                ColumnCount = table.ColumnCount;
            }

            internal bool FirstRow { get; }
            internal bool LastRow { get; }
            internal bool FirstColumn { get; }
            internal bool LastColumn { get; }
            internal bool BandedRows { get; }
            internal bool BandedColumns { get; }
            internal int RowCount { get; }
            internal int ColumnCount { get; }
        }

        // While a slide renders, each table's style options are read once rather than for every cell they apply to.
        [ThreadStatic] private static PowerPointTable? _styleFlagsTable;
        [ThreadStatic] private static TableStyleFlags _styleFlags;

        private static TableStyleFlags StyleFlags(PowerPointTable table) {
            if (_renderingSlide == null) return new TableStyleFlags(table);
            if (!ReferenceEquals(table, _styleFlagsTable)) {
                _styleFlags = new TableStyleFlags(table);
                _styleFlagsTable = table;
            }
            return _styleFlags;
        }
    }
}
