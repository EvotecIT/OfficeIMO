using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private const int MaximumTableRasterRows = 4096;
        private const int MaximumTableRasterColumns = 4096;
        private const long MaximumTableRasterCells = 100_000L;

        internal static bool TryCreateTableDrawing(PowerPointTable table,
            out OfficeDrawing drawing, out string? reason) {
            ArgumentNullException.ThrowIfNull(table);
            drawing = null!;
            if (!table.TryGetBoundsPoints(out double left, out double top,
                    out double width, out double height)
                || width <= 0D || height <= 0D) {
                reason = "The table has no positive renderable bounds.";
                return false;
            }
            if (!TryValidateTableRasterLimits(table, out reason)) {
                return false;
            }
            try {
                drawing = new OfficeDrawing(width, height);
                var diagnostics = new List<OfficeImageExportDiagnostic>();
                var mapping = new PowerPointShapeBoundsMapping(-left, -top,
                    1D, 1D);
                OfficeXmlNode colorScheme = table.OwnerSlide == null
                    ? default
                    : GetSlideColorScheme(table.OwnerSlide);
                AddTable(drawing, table, diagnostics, mapping, colorScheme);
                OfficeImageExportDiagnostic? failure = diagnostics
                    .FirstOrDefault(diagnostic => diagnostic.Severity
                        != OfficeImageExportDiagnosticSeverity.Info);
                if (failure != null || drawing.Elements.Count == 0) {
                    reason = failure?.Message
                        ?? "The table renderer produced no visible drawing content.";
                    drawing = null!;
                    return false;
                }
                reason = null;
                return true;
            } catch (Exception exception) when (exception is ArgumentException
                                                or InvalidOperationException
                                                or OverflowException) {
                drawing = null!;
                reason = $"The table cannot be projected through the shared drawing renderer: {exception.Message}";
                return false;
            }
        }

        private static OfficeColor ResolveTableCellFillColor(PowerPointTable table, PowerPointTableCell cell, int row, int column, OfficeXmlNode tableStyle, OfficeXmlNode colorScheme) =>
            cell.NoFill ? OfficeColor.Transparent
                : OfficeXmlThemeColors.ResolveColor(cell.Element.Child(D, "tcPr").Child(D, "solidFill"), colorScheme)
                ?? ResolveTableStyleFillColor(table, row, column, tableStyle, colorScheme)
                ?? OfficeColor.White;

        internal static OfficeColor ResolveTableCellFillColorForExport(
            PowerPointTable table, int row, int column) =>
            ResolveTableCellFillColorForExport(table, row, column, ResolveTableStyleElement(table));

        internal static OfficeColor ResolveTableCellFillColorForExport(
            PowerPointTable table, int row, int column, OfficeXmlNode tableStyle) {
            OfficeXmlNode colorScheme = ResolveTableColorScheme(table);
            return ResolveTableCellFillColor(table,
                table.GetCell(row, column), row, column,
                tableStyle, colorScheme);
        }

        internal static OfficeBorderBox ResolveTableCellBordersForExport(
            PowerPointTable table, int row, int column) =>
            ResolveTableCellBordersForExport(table, row, column, ResolveTableStyleElement(table));

        internal static OfficeBorderBox ResolveTableCellBordersForExport(
            PowerPointTable table, int row, int column, OfficeXmlNode tableStyle) {
            OfficeXmlNode colorScheme = ResolveTableColorScheme(table);
            PowerPointTableCell cell = table.GetCell(row, column);
            (int rowSpan, int columnSpan) = cell.Merge;
            return ResolveTableCellBorders(table,
                cell, row, column, Math.Max(1, rowSpan), Math.Max(1, columnSpan),
                tableStyle, colorScheme);
        }

        internal static OfficeColor ResolveTableCellFillColorForBinary(
            PowerPointTable table, int row, int column) =>
            ResolveTableCellFillColorForExport(table, row, column);

        internal static OfficeBorderBox ResolveTableCellBordersForBinary(
            PowerPointTable table, int row, int column) =>
            ResolveTableCellBordersForExport(table, row, column);

        private static OfficeBorderBox ResolveTableCellBorders(PowerPointTable table, PowerPointTableCell cell, int row, int column, int rowSpan, int columnSpan, OfficeXmlNode tableStyle, OfficeXmlNode colorScheme) {
            OfficeXmlNode properties = cell.Element.Child(D, "tcPr");
            return new OfficeBorderBox(
                ResolveTableCellBorderSide(properties.Child(D, "lnL"), colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Left, colorScheme)),
                ResolveTableCellBorderSide(properties.Child(D, "lnT"), colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Top, colorScheme)),
                ResolveTableCellBorderSide(properties.Child(D, "lnR"), colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Right, colorScheme)),
                ResolveTableCellBorderSide(properties.Child(D, "lnB"), colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Bottom, colorScheme)),
                ResolveOptionalTableCellBorderSide(properties.Child(D, "lnTlToBr"), colorScheme),
                ResolveOptionalTableCellBorderSide(properties.Child(D, "lnBlToTr"), colorScheme));
        }

        private static OfficeBorderSide ResolveTableCellBorderSide(OfficeXmlNode line, OfficeXmlNode colorScheme, OfficeBorderSide? fallback = null) {
            if (line.IsNone) {
                return fallback ?? new OfficeBorderSide(OfficeColor.FromRgb(191, 191, 191), 0.5D);
            }

            if (!line.Child(D, "noFill").IsNone) {
                return new OfficeBorderSide(OfficeColor.FromRgba(0, 0, 0, 0), 0D);
            }

            int? lineWidth = line.Int32Attribute("w");
            double width = lineWidth > 0
                ? lineWidth.Value / 12700D
                : 0.75D;
            OfficeColor color = OfficeXmlThemeColors.ResolveColor(line.Child(D, "solidFill"), colorScheme)
                ?? OfficeColor.FromRgb(191, 191, 191);
            string? dash = line.Child(D, "prstDash").Attribute("val");
            // DrawingML's double-line width includes both strokes and the intervening gap.
            if (line.Attribute("cmpd") == "dbl") {
                return new OfficeBorderSide(color, width / 3D, MapDash(dash),
                    OfficeBorderLineKind.Double, 2D * width / 3D);
            }
            return new OfficeBorderSide(color, width, MapDash(dash));
        }

        private static OfficeBorderSide? ResolveOptionalTableCellBorderSide(OfficeXmlNode line, OfficeXmlNode colorScheme) =>
            line.IsNone ? null : ResolveTableCellBorderSide(line, colorScheme);

    }
}
