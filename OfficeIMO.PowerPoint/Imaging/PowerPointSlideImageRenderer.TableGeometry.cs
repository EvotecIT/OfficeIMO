using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
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
}
