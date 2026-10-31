using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private const int MaximumTableRasterRows = 4096;
        private const int MaximumTableRasterColumns = 4096;
        private const long MaximumTableRasterCells = 100_000L;

        internal static bool TryCreateTableDrawing(PowerPointTable table,
            out OfficeDrawing drawing, out string? reason) {
            if (table == null) throw new ArgumentNullException(nameof(table));
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
                A.ColorScheme? colorScheme = table.OwnerSlide == null
                    ? null
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

        private static OfficeColor ResolveTableCellFillColor(PowerPointTable table, PowerPointTableCell cell, int row, int column, A.TableStyleEntry? tableStyle, A.ColorScheme? colorScheme) =>
            cell.NoFill ? OfficeColor.Transparent
                : OfficeOpenXmlThemeColorResolver.ResolveColor(cell.Cell.TableCellProperties?.GetFirstChild<A.SolidFill>(), colorScheme)
                ?? ResolveTableStyleFillColor(table, row, column, tableStyle, colorScheme)
                ?? OfficeColor.White;

        internal static OfficeColor ResolveTableCellFillColorForExport(
            PowerPointTable table, int row, int column) {
            A.ColorScheme? colorScheme = table.OwnerSlide == null
                ? PowerPointEffectiveRunStyleResolver.ResolveColorScheme(table.RelationshipOwnerPart)
                : GetSlideColorScheme(table.OwnerSlide);
            return ResolveTableCellFillColor(table,
                table.GetCell(row, column), row, column,
                ResolveTableStyle(table), colorScheme);
        }

        internal static OfficeBorderBox ResolveTableCellBordersForExport(
            PowerPointTable table, int row, int column) {
            A.ColorScheme? colorScheme = table.OwnerSlide == null
                ? PowerPointEffectiveRunStyleResolver.ResolveColorScheme(table.RelationshipOwnerPart)
                : GetSlideColorScheme(table.OwnerSlide);
            PowerPointTableCell cell = table.GetCell(row, column);
            (int rowSpan, int columnSpan) = cell.Merge;
            return ResolveTableCellBorders(table,
                cell, row, column, Math.Max(1, rowSpan), Math.Max(1, columnSpan),
                ResolveTableStyle(table), colorScheme);
        }

        internal static OfficeColor ResolveTableCellFillColorForBinary(
            PowerPointTable table, int row, int column) =>
            ResolveTableCellFillColorForExport(table, row, column);

        internal static OfficeBorderBox ResolveTableCellBordersForBinary(
            PowerPointTable table, int row, int column) =>
            ResolveTableCellBordersForExport(table, row, column);

        private static OfficeBorderBox ResolveTableCellBorders(PowerPointTable table, PowerPointTableCell cell, int row, int column, int rowSpan, int columnSpan, A.TableStyleEntry? tableStyle, A.ColorScheme? colorScheme) {
            A.TableCellProperties? properties = cell.Cell.TableCellProperties;
            return new OfficeBorderBox(
                ResolveTableCellBorderSide(properties?.LeftBorderLineProperties, colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Left, colorScheme)),
                ResolveTableCellBorderSide(properties?.TopBorderLineProperties, colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Top, colorScheme)),
                ResolveTableCellBorderSide(properties?.RightBorderLineProperties, colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Right, colorScheme)),
                ResolveTableCellBorderSide(properties?.BottomBorderLineProperties, colorScheme, ResolveTableStyleBorderSide(table, row, column, rowSpan, columnSpan, tableStyle, TableCellBorderEdge.Bottom, colorScheme)),
                ResolveOptionalTableCellBorderSide(properties?.TopLeftToBottomRightBorderLineProperties, colorScheme),
                ResolveOptionalTableCellBorderSide(properties?.BottomLeftToTopRightBorderLineProperties, colorScheme));
        }

        private static OfficeBorderSide ResolveTableCellBorderSide(A.LinePropertiesType? line, A.ColorScheme? colorScheme, OfficeBorderSide? fallback = null) {
            if (line == null) {
                return fallback ?? new OfficeBorderSide(OfficeColor.FromRgb(191, 191, 191), 0.5D);
            }

            if (line.GetFirstChild<A.NoFill>() != null) {
                return new OfficeBorderSide(OfficeColor.FromRgba(0, 0, 0, 0), 0D);
            }

            double width = line?.Width?.Value > 0
                ? line.Width.Value / 12700D
                : 0.75D;
            OfficeColor color = OfficeOpenXmlThemeColorResolver.ResolveColor(line?.GetFirstChild<A.SolidFill>(), colorScheme)
                ?? OfficeColor.FromRgb(191, 191, 191);
            A.PresetDash? dash = line?.GetFirstChild<A.PresetDash>();
            return new OfficeBorderSide(color, width, MapDash(dash?.Val?.Value));
        }

        private static OfficeBorderSide? ResolveOptionalTableCellBorderSide(A.LinePropertiesType? line, A.ColorScheme? colorScheme) =>
            line == null ? null : ResolveTableCellBorderSide(line, colorScheme);

    }
}
