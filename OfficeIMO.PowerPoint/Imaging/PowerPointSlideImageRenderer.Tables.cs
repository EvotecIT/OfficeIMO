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

        private static void AddTable(OfficeDrawing drawing, PowerPointTable table, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
            if (!TryValidateTableRasterLimits(table, out string? limitReason)) {
                AddUnsupportedShapeDiagnostic(diagnostics, table,
                    limitReason!);
                return;
            }
            if (!TryGetBounds(table, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            if (table.Rows == 0 || table.Columns == 0) {
                AddUnsupportedShapeDiagnostic(diagnostics, table, "Skipped an empty PowerPoint table.");
                return;
            }

            double[] columnWidths = CreateTableColumnWidths(table, width);
            double[] rowHeights = CreateTableRowHeights(table, height);
            double[] columnLefts = CreateOffsets(columnWidths);
            double[] rowTops = CreateOffsets(rowHeights);
            A.TableStyleEntry? tableStyle = ResolveTableStyle(table);
            bool[,] coveredCells = new bool[table.Rows, table.Columns];

            for (int row = 0; row < table.Rows; row++) {
                for (int column = 0; column < table.Columns; column++) {
                    if (coveredCells[row, column]) {
                        continue;
                    }

                    PowerPointTableCell cell = table.GetCell(row, column);
                    if (cell.IsMergedCell) {
                        continue;
                    }

                    (int rowSpan, int columnSpan) = cell.Merge;
                    MarkCoveredTableCells(coveredCells, row, column, rowSpan, columnSpan);
                    double cellLeft = left + columnLefts[column];
                    double cellTop = top + rowTops[row];
                    double cellWidth = SumSpan(columnWidths, column, columnSpan);
                    double cellHeight = SumSpan(rowHeights, row, rowSpan);
                    if (cellWidth <= 0D || cellHeight <= 0D) {
                        continue;
                    }

                    AddTableCellShape(drawing, table, cell, row, column, rowSpan, columnSpan, tableStyle, cellLeft, cellTop, cellWidth, cellHeight, left, top, width, height, colorScheme);
                    AddTableCellText(drawing, table, cell, cellLeft, cellTop, cellWidth, cellHeight, left, top, width, height, diagnostics, mapping, colorScheme);
                }
            }
        }

        private static bool TryValidateTableRasterLimits(
            PowerPointTable table, out string? reason) {
            long cellCount = checked((long)table.Rows * table.Columns);
            if (table.Rows > MaximumTableRasterRows
                || table.Columns > MaximumTableRasterColumns
                || cellCount > MaximumTableRasterCells) {
                reason = $"The table rasterization request ({table.Rows} rows, {table.Columns} columns, {cellCount} cells) exceeds the safe limit of {MaximumTableRasterRows} rows, {MaximumTableRasterColumns} columns, or {MaximumTableRasterCells} cells.";
                return false;
            }
            reason = null;
            return true;
        }

        private static void MarkCoveredTableCells(bool[,] coveredCells, int row, int column, int rowSpan, int columnSpan) {
            int rowEnd = Math.Min(coveredCells.GetLength(0), row + Math.Max(1, rowSpan));
            int columnEnd = Math.Min(coveredCells.GetLength(1), column + Math.Max(1, columnSpan));
            for (int coveredRow = row; coveredRow < rowEnd; coveredRow++) {
                for (int coveredColumn = column; coveredColumn < columnEnd; coveredColumn++) {
                    if (coveredRow == row && coveredColumn == column) {
                        continue;
                    }

                    coveredCells[coveredRow, coveredColumn] = true;
                }
            }
        }

        private static void AddTableCellShape(
            OfficeDrawing drawing,
            PowerPointTable table,
            PowerPointTableCell cell,
            int row,
            int column,
            int rowSpan,
            int columnSpan,
            A.TableStyleEntry? tableStyle,
            double cellLeft,
            double cellTop,
            double cellWidth,
            double cellHeight,
            double tableLeft,
            double tableTop,
            double tableWidth,
            double tableHeight,
            A.ColorScheme? colorScheme) {
            OfficeColor fillColor = ResolveTableCellFillColor(table, cell, row, column, tableStyle, colorScheme);

            OfficeTransform? transform = CreateTableCellTransform(table, cellLeft, cellTop, tableLeft, tableTop, tableWidth, tableHeight);
            drawing.AddBorderBox(cellLeft, cellTop, cellWidth, cellHeight, fillColor, ResolveTableCellBorders(table, cell, row, column, rowSpan, columnSpan, tableStyle, colorScheme), transform);
        }

        private static void AddTableCellText(
            OfficeDrawing drawing,
            PowerPointTable table,
            PowerPointTableCell cell,
            double cellLeft,
            double cellTop,
            double cellWidth,
            double cellHeight,
            double tableLeft,
            double tableTop,
            double tableWidth,
            double tableHeight,
            List<OfficeImageExportDiagnostic> diagnostics,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme) {
            string text = cell.Text;
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            double marginLeft = mapping.MapHorizontalLength(cell.PaddingLeftPoints ?? 3.6D);
            double marginTop = mapping.MapVerticalLength(cell.PaddingTopPoints ?? 1.8D);
            double marginRight = mapping.MapHorizontalLength(cell.PaddingRightPoints ?? 3.6D);
            double marginBottom = mapping.MapVerticalLength(cell.PaddingBottomPoints ?? 1.8D);
            double textWidth = cellWidth - marginLeft - marginRight;
            double textHeight = cellHeight - marginTop - marginBottom;
            if (textWidth <= 0D || textHeight <= 0D) {
                AddUnsupportedShapeDiagnostic(diagnostics, table, "Skipped PowerPoint table cell text because the cell margins leave no renderable drawing area.");
                return;
            }

            if (TryAddTableCellParagraphFlow(
                drawing,
                table,
                cell,
                cellLeft,
                cellTop,
                cellWidth,
                cellHeight,
                textWidth,
                textHeight,
                marginLeft,
                marginTop,
                tableLeft,
                tableTop,
                tableWidth,
                tableHeight,
                mapping,
                colorScheme,
                diagnostics)) {
                return;
            }

            var padding = new OfficeTextPadding(marginLeft, marginTop, marginRight, marginBottom);
            List<OfficeRichTextRun> richRuns = CreateRichTextRuns(cell, colorScheme, mapping);
            if (ShouldRenderRichText(richRuns)) {
                drawing.AddRichText(
                    richRuns,
                    cellLeft,
                    cellTop,
                    cellWidth,
                    cellHeight,
                    MapTextAlignment(cell.HorizontalAlignment),
                    verticalAlignment: MapTextVerticalAlignment(cell.VerticalAlignment),
                    rotationDegrees: table.Rotation ?? 0D,
                    rotationCenterX: tableLeft + (tableWidth / 2D),
                    rotationCenterY: tableTop + (tableHeight / 2D),
                    wrapText: true,
                    flipHorizontal: table.HorizontalFlip == true,
                    flipVertical: table.VerticalFlip == true,
                    padding: padding);
                return;
            }

            drawing.AddText(
                text,
                cellLeft,
                cellTop,
                cellWidth,
                cellHeight,
                CreateFont(cell, mapping),
                ResolveTableCellTextColor(cell, colorScheme),
                MapTextAlignment(cell.HorizontalAlignment),
                verticalAlignment: MapTextVerticalAlignment(cell.VerticalAlignment),
                rotationDegrees: table.Rotation ?? 0D,
                rotationCenterX: tableLeft + (tableWidth / 2D),
                rotationCenterY: tableTop + (tableHeight / 2D),
                wrapText: true,
                flipHorizontal: table.HorizontalFlip == true,
                flipVertical: table.VerticalFlip == true,
                padding: padding);
        }

        private static double[] CreateTableColumnWidths(PowerPointTable table, double tableWidth) {
            double[] widths = new double[table.Columns];
            double fallback = table.Columns > 0 ? tableWidth / table.Columns : tableWidth;
            for (int column = 0; column < widths.Length; column++) {
                double width = table.GetColumnWidthPoints(column);
                widths[column] = width > 0D ? width : fallback;
            }

            NormalizeTableSegments(widths, tableWidth);
            return widths;
        }

        private static double[] CreateTableRowHeights(PowerPointTable table, double tableHeight) {
            double[] heights = new double[table.Rows];
            double fallback = table.Rows > 0 ? tableHeight / table.Rows : tableHeight;
            for (int row = 0; row < heights.Length; row++) {
                double height = table.GetRowHeightPoints(row);
                heights[row] = height > 0D ? height : fallback;
            }

            NormalizeTableSegments(heights, tableHeight);
            return heights;
        }

        private static void NormalizeTableSegments(double[] values, double targetTotal) {
            double currentTotal = 0D;
            for (int i = 0; i < values.Length; i++) {
                currentTotal += values[i];
            }

            if (currentTotal <= 0D) {
                double fallback = values.Length > 0 ? targetTotal / values.Length : targetTotal;
                for (int i = 0; i < values.Length; i++) {
                    values[i] = fallback;
                }

                return;
            }

            double ratio = targetTotal / currentTotal;
            for (int i = 0; i < values.Length; i++) {
                values[i] *= ratio;
            }
        }

        private static double[] CreateOffsets(double[] values) {
            double[] offsets = new double[values.Length];
            double current = 0D;
            for (int i = 0; i < values.Length; i++) {
                offsets[i] = current;
                current += values[i];
            }

            return offsets;
        }

        private static double SumSpan(double[] values, int start, int span) {
            int count = Math.Max(1, span);
            int end = Math.Min(values.Length, start + count);
            double total = 0D;
            for (int i = start; i < end; i++) {
                total += values[i];
            }

            return total;
        }

        private static OfficeTransform? CreateTableCellTransform(
            PowerPointTable table,
            double cellLeft,
            double cellTop,
            double tableLeft,
            double tableTop,
            double tableWidth,
            double tableHeight) {
            double rotation = table.Rotation ?? 0D;
            bool flipHorizontal = table.HorizontalFlip == true;
            bool flipVertical = table.VerticalFlip == true;
            if (Math.Abs(rotation) < 0.000001D && !flipHorizontal && !flipVertical) {
                return null;
            }

            var transform = new OfficeImageFrameTransform(
                rotation,
                tableLeft + (tableWidth / 2D) - cellLeft,
                tableTop + (tableHeight / 2D) - cellTop,
                flipHorizontal,
                flipVertical);
            return transform.CreateDestinationTransform();
        }

        private static OfficeColor ResolveTableCellFillColor(PowerPointTable table, PowerPointTableCell cell, int row, int column, A.TableStyleEntry? tableStyle, A.ColorScheme? colorScheme) =>
            OfficeOpenXmlThemeColorResolver.ResolveColor(cell.Cell.TableCellProperties?.GetFirstChild<A.SolidFill>(), colorScheme)
                ?? ResolveTableStyleFillColor(table, row, column, tableStyle, colorScheme)
                ?? OfficeColor.White;

        internal static OfficeColor ResolveTableCellFillColorForExport(
            PowerPointTable table, int row, int column) {
            A.ColorScheme? colorScheme = table.OwnerSlide == null
                ? null
                : GetSlideColorScheme(table.OwnerSlide);
            return ResolveTableCellFillColor(table,
                table.GetCell(row, column), row, column,
                ResolveTableStyle(table), colorScheme);
        }

        internal static OfficeBorderBox ResolveTableCellBordersForExport(
            PowerPointTable table, int row, int column) {
            A.ColorScheme? colorScheme = table.OwnerSlide == null
                ? null
                : GetSlideColorScheme(table.OwnerSlide);
            return ResolveTableCellBorders(table,
                table.GetCell(row, column), row, column, 1, 1,
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

        private static OfficeFontInfo CreateFont(PowerPointTableCell cell, PowerPointShapeBoundsMapping mapping) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (cell.Bold) {
                style |= OfficeFontStyle.Bold;
            }

            if (cell.Italic) {
                style |= OfficeFontStyle.Italic;
            }

            return new OfficeFontInfo(cell.FontName ?? "Calibri", mapping.MapFontSize(cell.FontSize ?? 10), style);
        }

        private static List<OfficeRichTextRun> CreateRichTextRuns(PowerPointTableCell cell, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
            var richRuns = new List<OfficeRichTextRun>();
            A.TextBody? textBody = cell.Cell.TextBody;
            if (textBody == null) {
                return richRuns;
            }

            foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
                List<A.Run> paragraphRuns = paragraph.Elements<A.Run>().ToList();
                bool paragraphHasVisibleText = paragraphRuns.Any(run => !string.IsNullOrEmpty(run.Text?.Text));
                if (!paragraphHasVisibleText) {
                    continue;
                }

                if (richRuns.Count > 0) {
                    richRuns.Add(CreateRichTextRun(Environment.NewLine, null, cell, colorScheme, mapping));
                }

                for (int runIndex = 0; runIndex < paragraphRuns.Count; runIndex++) {
                    A.Run sourceRun = paragraphRuns[runIndex];
                    string runText = sourceRun.Text?.Text ?? string.Empty;
                    if (string.IsNullOrEmpty(runText)) {
                        continue;
                    }

                    richRuns.Add(CreateRichTextRun(runText, new PowerPointTextRun(sourceRun), cell, colorScheme, mapping));
                }
            }

            return richRuns;
        }

        private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTableCell cell, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
            return CreateRichTextRun(text, run, cell, paragraph: null, colorScheme, mapping, markerRun: false);
        }

        private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTableCell cell, PowerPointParagraph? paragraph, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, bool markerRun = false) {
            OfficeColor color = ResolveTableCellTextRunColor(run, cell, colorScheme);
            OfficeColor? backgroundColor = ResolveTableCellTextRunBackgroundColor(run, colorScheme);
            return new OfficeRichTextRun(
                text,
                mapping.MapFontSize(markerRun ? paragraph?.BulletSizePoints ?? run?.FontSize ?? cell.FontSize ?? 10 : run?.FontSize ?? cell.FontSize ?? 10),
                color,
                run?.Bold == true,
                run?.Italic == true,
                run?.Underline == true,
                markerRun ? paragraph?.BulletFontName ?? run?.FontName ?? cell.FontName ?? "Calibri" : run?.FontName ?? cell.FontName ?? "Calibri",
                run?.Strikethrough == true,
                backgroundColor);
        }

        private static OfficeColor ResolveTableCellTextRunColor(PowerPointTextRun? run, PowerPointTableCell cell, A.ColorScheme? colorScheme) {
            OfficeColor? runColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run?.Run.RunProperties?.GetFirstChild<A.SolidFill>(), colorScheme);
            if (runColor.HasValue) {
                return runColor.Value;
            }

            return ResolveTableCellTextColor(cell, colorScheme);
        }

        private static OfficeColor ResolveTableCellTextColor(PowerPointTableCell cell, A.ColorScheme? colorScheme) {
            A.Run? run = cell.Cell.TextBody?
                .Elements<A.Paragraph>()
                .SelectMany(paragraph => paragraph.Elements<A.Run>())
                .FirstOrDefault();
            OfficeColor? cellColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run?.RunProperties?.GetFirstChild<A.SolidFill>(), colorScheme);
            return cellColor.HasValue
                ? cellColor.Value
                : OfficeColor.Black;
        }

        private static OfficeColor? ResolveTableCellTextRunBackgroundColor(PowerPointTextRun? run, A.ColorScheme? colorScheme) {
            return OfficeOpenXmlThemeColorResolver.ResolveColor(run?.Run.RunProperties?.GetFirstChild<A.Highlight>(), colorScheme);
        }
    }
}
