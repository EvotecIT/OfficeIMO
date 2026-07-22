using AngleSharp.Dom;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static double ImportTable(
        IElement tableElement,
        PptCore.PowerPointSlide slide,
        double top,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        HtmlSemanticBlock? semanticBlock = null) {
        PowerPointHtmlTableGrid grid = BuildTableGrid(tableElement, budget, result);
        if (grid.Rows == 0 || grid.Columns == 0) {
            return top;
        }

        if (!budget.TryReserveTableWithShape(out string tableLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "A slide table was omitted because the shared import limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, detail: tableLimit);
            return top;
        }

        double fallbackWidth = Math.Max(240D, grid.Columns * 150D);
        double fallbackHeight = Math.Max(70D, grid.Rows * 34D);
        ReadSemanticShapeGeometry(tableElement, 64D, top, fallbackWidth, fallbackHeight, budget, result,
            out double left, out double tableTop, out double width, out double height);
        PptCore.PowerPointTable table = slide.AddTablePoints(grid.Rows, grid.Columns, left, tableTop, width, height);
        foreach (PowerPointHtmlTableCell cell in grid.Cells) {
            table.GetCell(cell.Row, cell.Column).Text = cell.Text;
            if (cell.RowSpan > 1 || cell.ColumnSpan > 1) {
                table.MergeCells(cell.Row, cell.Column, cell.Row + cell.RowSpan - 1, cell.Column + cell.ColumnSpan - 1);
                result.MergedRanges++;
            }
        }

        if (semanticBlock?.Table != null) ApplySemanticTableFormatting(table, semanticBlock.Table);

        ApplyShapeTransforms(tableElement, table, budget, result);
        result.Tables++;
        return Math.Max(top + Math.Max(90D, grid.Rows * 40D), tableTop + height + 20D);
    }

    private static void ApplySemanticTableFormatting(PptCore.PowerPointTable target, HtmlSemanticTable source) {
        var occupied = new HashSet<long>();
        int rowIndex = 0;
        foreach (HtmlSemanticTableRow row in source.Rows) {
            int columnIndex = 0;
            foreach (HtmlSemanticTableCell cell in row.Cells) {
                while (occupied.Contains(GetTableCellKey(rowIndex, columnIndex))) columnIndex++;
                if (rowIndex >= target.Rows || columnIndex >= target.Columns) break;
                PptCore.PowerPointTableCell targetCell = target.GetCell(rowIndex, columnIndex);
                if (cell.Runs.Count > 0) ApplySemanticRuns(targetCell.Paragraphs[0], cell.Runs);
                if (cell.IsHeader) {
                    foreach (PptCore.PowerPointTextRun run in targetCell.Runs) run.Bold = true;
                }
                string fill = NormalizeSemanticColor(cell.Style?.GetValue("background-color"));
                if (fill.Length > 0) targetCell.FillColor = fill;
                string color = NormalizeSemanticColor(cell.Style?.GetValue("color"));
                if (color.Length > 0) {
                    foreach (PptCore.PowerPointTextRun run in targetCell.Runs) run.Color = color;
                }
                ApplySemanticTableAlignment(targetCell, cell.Style?.GetValue("text-align"));
                ReservePowerPointSpan(occupied, rowIndex, columnIndex,
                    Math.Max(1, cell.RowSpan), Math.Max(1, cell.ColumnSpan));
                columnIndex += Math.Max(1, cell.ColumnSpan);
            }
            rowIndex++;
            if (rowIndex >= target.Rows) break;
        }
    }

    private static void ApplySemanticTableAlignment(PptCore.PowerPointTableCell cell, string? alignment) {
        switch ((alignment ?? string.Empty).Trim().ToLowerInvariant()) {
            case "center":
                cell.HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center;
                break;
            case "right":
            case "end":
                cell.HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right;
                break;
            case "justify":
                cell.HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Justified;
                break;
            case "left":
            case "start":
                cell.HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left;
                break;
        }
    }

    private static PowerPointHtmlTableGrid BuildTableGrid(IElement table, HtmlImportBudget budget, HtmlToPowerPointResult result) {
        int maxTableCells = budget.Limits.MaxTableCells;
        var cells = new List<PowerPointHtmlTableCell>();
        var occupied = new HashSet<long>();
        int rowIndex = 0;
        int rowExtent = 0;
        int columnExtent = 0;

        foreach (IElement row in EnumerateDirectTableRows(table)) {
            int columnIndex = 0;
            foreach (IElement element in row.Children.Where(IsPowerPointTableCell)) {
                while (occupied.Contains(GetTableCellKey(rowIndex, columnIndex))) {
                    columnIndex++;
                }

                int rowSpan = ReadPowerPointSpan(element, "rowspan", result);
                int columnSpan = ReadPowerPointSpan(element, "colspan", result);
                if ((long)rowSpan * columnSpan > maxTableCells) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An HTML table span exceeded the configured MaxTableCells limit; the span was ignored.", lossKind: HtmlConversionLossKind.Approximation);
                    rowSpan = 1;
                    columnSpan = 1;
                }

                long candidateRowsLong = Math.Max((long)rowExtent, (long)rowIndex + rowSpan);
                long candidateColumnsLong = Math.Max((long)columnExtent, (long)columnIndex + columnSpan);
                if (candidateRowsLong * candidateColumnsLong > maxTableCells) {
                    rowSpan = 1;
                    columnSpan = 1;
                    int candidateRows = Math.Max(rowExtent, rowIndex + 1);
                    int candidateColumns = Math.Max(columnExtent, columnIndex + 1);
                    if ((long)candidateRows * candidateColumns > maxTableCells) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                            "HTML table exceeded the configured MaxTableCells limit; remaining cells were skipped.", lossKind: HtmlConversionLossKind.Omission);
                        return new PowerPointHtmlTableGrid(rowExtent, columnExtent, cells);
                    }

                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An HTML table span exceeded the configured MaxTableCells limit; the span was ignored.", lossKind: HtmlConversionLossKind.Approximation);
                }

                if (PowerPointSpanOverlaps(occupied, rowIndex, columnIndex, rowSpan, columnSpan)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                        "An HTML table cell contained an overlapping span; the span was ignored.", lossKind: HtmlConversionLossKind.Approximation);
                    rowSpan = 1;
                    columnSpan = 1;
                }

                ReservePowerPointSpan(occupied, rowIndex, columnIndex, rowSpan, columnSpan);
                string text = PreserveText(element.TextContent);
                if (!budget.IsMetadataWithinLimit(text, out string metadataLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                        "A slide table cell was imported without text because it exceeded the shared field limit.",
                        lossKind: HtmlConversionLossKind.Omission, detail: metadataLimit);
                    text = string.Empty;
                }
                cells.Add(new PowerPointHtmlTableCell(
                    rowIndex,
                    columnIndex,
                    rowSpan,
                    columnSpan,
                    text));
                rowExtent = Math.Max(rowExtent, rowIndex + rowSpan);
                columnExtent = Math.Max(columnExtent, columnIndex + columnSpan);
                columnIndex += columnSpan;
            }

            rowIndex++;
            rowExtent = Math.Max(rowExtent, rowIndex);
            if ((long)Math.Max(1, rowExtent) * Math.Max(1, columnExtent) > maxTableCells) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "HTML table exceeded the configured MaxTableCells limit; remaining rows were skipped.", lossKind: HtmlConversionLossKind.Omission);
                break;
            }
        }

        return new PowerPointHtmlTableGrid(rowExtent, columnExtent, cells);
    }

    private static IEnumerable<IElement> EnumerateDirectTableRows(IElement table) {
        foreach (IElement child in table.Children) {
            if (IsElement(child, "tr")) {
                yield return child;
            } else if (IsElement(child, "thead") || IsElement(child, "tbody") || IsElement(child, "tfoot")) {
                foreach (IElement row in child.Children.Where(element => IsElement(element, "tr"))) {
                    yield return row;
                }
            }
        }
    }

    private static bool IsPowerPointTableCell(IElement element) => IsElement(element, "th") || IsElement(element, "td");

    private static int ReadPowerPointSpan(IElement cell, string attributeName, HtmlToPowerPointResult result) {
        string? raw = cell.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return 1;
        }

        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span) || span <= 0) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                "An HTML table cell contained an invalid " + attributeName + " value; a span of 1 was used.", lossKind: HtmlConversionLossKind.Approximation);
            return 1;
        }

        return span;
    }

    private static bool PowerPointSpanOverlaps(HashSet<long> occupied, int row, int column, int rowSpan, int columnSpan) {
        for (int currentRow = row; currentRow < row + rowSpan; currentRow++) {
            for (int currentColumn = column; currentColumn < column + columnSpan; currentColumn++) {
                if (occupied.Contains(GetTableCellKey(currentRow, currentColumn))) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void ReservePowerPointSpan(HashSet<long> occupied, int row, int column, int rowSpan, int columnSpan) {
        for (int currentRow = row; currentRow < row + rowSpan; currentRow++) {
            for (int currentColumn = column; currentColumn < column + columnSpan; currentColumn++) {
                occupied.Add(GetTableCellKey(currentRow, currentColumn));
            }
        }
    }

    private static long GetTableCellKey(int row, int column) => ((long)row << 32) | (uint)column;

    private sealed class PowerPointHtmlTableGrid {
        internal PowerPointHtmlTableGrid(int rows, int columns, IReadOnlyList<PowerPointHtmlTableCell> cells) {
            Rows = rows;
            Columns = columns;
            Cells = cells;
        }

        internal int Rows { get; }
        internal int Columns { get; }
        internal IReadOnlyList<PowerPointHtmlTableCell> Cells { get; }
    }

    private sealed class PowerPointHtmlTableCell {
        internal PowerPointHtmlTableCell(int row, int column, int rowSpan, int columnSpan, string text) {
            Row = row;
            Column = column;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
            Text = text;
        }

        internal int Row { get; }
        internal int Column { get; }
        internal int RowSpan { get; }
        internal int ColumnSpan { get; }
        internal string Text { get; }
    }
}
