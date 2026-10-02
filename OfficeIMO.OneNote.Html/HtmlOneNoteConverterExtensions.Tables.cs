using OfficeIMO.Html;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.OneNote.Html;

/// <summary>Native rectangular table projection and bounded column sizing.</summary>
public static partial class HtmlOneNoteConverterExtensions {
    private static void ImportTable(
        HtmlSemanticBlock source,
        IList<OneNoteElement> target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        HtmlSemanticTable? sourceTable = source.Table;
        if (sourceTable == null || sourceTable.Rows.Count == 0) return;

        if (!TryProjectTableGrid(sourceTable, budget.Limits.MaxTableCells,
            out List<List<HtmlSemanticTableCell?>> grid, out int columns,
            out bool approximatedGrid, out string gridLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "The HTML table could not fit OneNote's rectangular native table limits.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission, gridLimit);
            return;
        }
        if (!budget.TryReserveTableWithShape(out string tableLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An HTML table was omitted because the shared import limit was reached.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission,
                tableLimit);
            return;
        }
        var table = new OneNoteTable { BordersVisible = true };
        foreach (List<HtmlSemanticTableCell?> gridRow in grid) {
            var row = new OneNoteTableRow();
            for (int column = 0; column < columns; column++) {
                HtmlSemanticTableCell? cellElement = column < gridRow.Count ? gridRow[column] : null;
                var cell = new OneNoteTableCell();
                if (cellElement != null && TryParseArgb(cellElement.Style?.GetValue("background-color"), out uint shading)) {
                    cell.ShadingColorArgb = shading;
                }
                if (cellElement != null) {
                    OneNoteParagraph? paragraph = CreateParagraph(cellElement.Text, cellElement.Runs, 0, cellElement.Style, result, budget);
                    if (paragraph != null) cell.Content.Add(paragraph);
                }
                if (cellElement != null && options.ImportImages) {
                    foreach (HtmlSemanticResource resource in cellElement.Resources.Where(item => item.Kind == HtmlResourceKind.Image)) {
                        ImportImage(resource, cell.Content, result, budget);
                    }
                }
                row.Cells.Add(cell);
            }
            table.Rows.Add(row);
        }
        if (approximatedGrid) {
            Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "HTML spans or uneven rows were flattened to OneNote's rectangular editable table; merged-cell alignment is not retained.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation,
                "columns=" + columns);
        }
        SetImportedTableColumnWidths(source, table, grid, result);
        string caption = string.Concat(sourceTable.CaptionRuns.Select(run => run.Text));
        if (caption.Length > 0) {
            OneNoteParagraph? captionParagraph = CreateParagraph(caption, sourceTable.CaptionRuns,
                0, null, result, budget);
            if (captionParagraph != null) {
                target.Add(captionParagraph);
                Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                    "A table caption was retained as a preceding paragraph because OneNote tables have no native caption.",
                    HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation,
                    "projection=adjacentParagraph");
            }
        }
        target.Add(table);
        result.Elements++;
        result.Tables++;
    }

    private static bool TryProjectTableGrid(
        HtmlSemanticTable source, int maxTableCells,
        out List<List<HtmlSemanticTableCell?>> grid, out int columns,
        out bool approximated, out string limit) {
        grid = new List<List<HtmlSemanticTableCell?>>();
        columns = 1;
        approximated = false;
        limit = string.Empty;
        int rows = source.Rows.Count;
        if (rows > maxTableCells) {
            limit = "rows=" + rows + ";maxCells=" + maxTableCells;
            return false;
        }

        // Occupancy preserves the column of cells following a rowspan. Only the
        // anchor is materialized with content; covered positions stay editable blanks.
        var occupiedUntilRow = new int[byte.MaxValue];
        for (int column = 0; column < occupiedUntilRow.Length; column++) occupiedUntilRow[column] = -1;
        for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
            int sourceRowIndex = source.Rows[rowIndex].SourceRowIndex;
            var projected = new List<HtmlSemanticTableCell?>();
            for (int column = 0; column < occupiedUntilRow.Length; column++) {
                if (occupiedUntilRow[column] >= sourceRowIndex) {
                    while (projected.Count <= column) projected.Add(null);
                }
            }
            int cursor = 0;
            foreach (HtmlSemanticTableCell cell in source.Rows[rowIndex].Cells) {
                int span = Math.Max(1, cell.ColumnSpan);
                if (span > byte.MaxValue) {
                    limit = "columns=" + span + ";maxColumns=" + byte.MaxValue;
                    return false;
                }
                while (cursor + span <= byte.MaxValue) {
                    bool free = true;
                    for (int offset = 0; offset < span; offset++) {
                        if (occupiedUntilRow[cursor + offset] >= sourceRowIndex) {
                            free = false;
                            cursor += offset + 1;
                            break;
                        }
                    }
                    if (free) break;
                }
                if (cursor + span > byte.MaxValue) {
                    limit = "columns>" + byte.MaxValue + ";maxColumns=" + byte.MaxValue;
                    return false;
                }
                while (projected.Count < cursor + span) projected.Add(null);
                projected[cursor] = cell;
                int finalRow = (int)Math.Min(source.Rows[rows - 1].SourceRowIndex,
                    sourceRowIndex + (long)Math.Max(1, cell.RowSpan) - 1L);
                for (int offset = 0; offset < span; offset++) occupiedUntilRow[cursor + offset] = finalRow;
                approximated |= cell.RowSpan != 1 || cell.ColumnSpan != 1;
                cursor += span;
            }
            columns = Math.Max(columns, projected.Count);
            if ((long)columns * rows > maxTableCells) {
                limit = "columns=" + columns + ";rows=" + rows + ";maxCells=" + maxTableCells;
                return false;
            }
            grid.Add(projected);
        }
        int projectedColumns = columns;
        approximated |= grid.Any(row => row.Count < projectedColumns);
        return true;
    }

    private static void SetImportedTableColumnWidths(
        HtmlSemanticBlock source, OneNoteTable table,
        IReadOnlyList<List<HtmlSemanticTableCell?>> grid, HtmlToOneNoteSectionResult result) {
        int columns = table.Rows.Max(row => row.Cells.Count);
        if (columns == 0) return;

        // Native writing otherwise assigns one half-inch unit to every column,
        // making ordinary imported tables unreadably narrow after reopen.
        const double defaultWidthHalfInches = 15D;
        double totalWidth = defaultWidthHalfInches;
        string? authoredWidth = source.Style?.GetValue("width");
        if (source.Style?.IsSpecifiedValue("width") != true && !IsAuthoredWidth(authoredWidth)) {
            authoredWidth = source.SourceElement.GetAttribute("width");
            if (double.TryParse(authoredWidth, NumberStyles.Float, CultureInfo.InvariantCulture, out double pixels)
                && pixels > 0D && !double.IsInfinity(pixels)) authoredWidth += "px";
        }
        bool hasAuthoredTableWidth = TryParseTableWidth(authoredWidth, defaultWidthHalfInches, out double requestedWidth);
        if (hasAuthoredTableWidth) {
            totalWidth = Math.Min(defaultWidthHalfInches, Math.Max(columns, requestedWidth));
            if (totalWidth != requestedWidth) ReportTableWidthApproximation(result, authoredWidth);
        } else if (IsAuthoredWidth(authoredWidth)) {
            ReportTableWidthApproximation(result, authoredWidth);
        }
        totalWidth = Math.Max(columns, totalWidth);

        var weights = new double[columns];
        foreach (OneNoteTableRow row in table.Rows) {
            for (int column = 0; column < row.Cells.Count; column++) {
                int characters = 0;
                foreach (OneNoteTextRun run in row.Cells[column].Content.OfType<OneNoteParagraph>()
                    .SelectMany(paragraph => paragraph.Runs)) {
                    characters = Math.Min(256, characters + Math.Min(256, run.Text.Length));
                    if (characters == 256) break;
                }
                weights[column] = Math.Max(weights[column], Math.Sqrt(Math.Min(256, characters) + 1D));
            }
        }

        var widths = new double[columns];
        var reportedConflicts = new bool[columns];
        double explicitTotal = 0D;
        for (int row = 0; row < table.Rows.Count; row++) {
            for (int column = 0; column < grid[row].Count; column++) {
                HtmlSemanticTableCell? cell = grid[row][column];
                if (cell == null) continue;
                string? cellWidth = cell.Style?.GetValue("width");
                if (!IsAuthoredWidth(cellWidth)) continue;
                if (cell.ColumnSpan != 1 || !TryParseTableWidth(cellWidth, totalWidth, out double requestedColumnWidth)) {
                    ReportTableWidthApproximation(result, cellWidth);
                    continue;
                }
                if (widths[column] == 0D) {
                    widths[column] = Math.Max(1D, requestedColumnWidth);
                    explicitTotal += widths[column];
                    if (widths[column] != requestedColumnWidth) ReportTableWidthApproximation(result, cellWidth);
                } else if (!reportedConflicts[column]
                    && Math.Abs(widths[column] - requestedColumnWidth) > 0.01D) {
                    ReportTableWidthApproximation(result, cellWidth);
                    reportedConflicts[column] = true;
                }
            }
        }

        int inferredColumns = widths.Count(width => width == 0D);
        if (explicitTotal > totalWidth - inferredColumns) {
            Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Authored HTML table columns exceeded OneNote's available page width and were scaled.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation);
            double available = Math.Max(0D, totalWidth - columns);
            double excess = explicitTotal - (columns - inferredColumns);
            for (int column = 0; column < columns; column++) {
                if (widths[column] > 0D) widths[column] = 1D + (excess > 0D ? available * (widths[column] - 1D) / excess : 0D);
            }
            explicitTotal = widths.Sum();
        } else if (hasAuthoredTableWidth && inferredColumns == 0 && explicitTotal < totalWidth) {
            Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Authored HTML table columns left unused table width, which was distributed proportionally.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation);
            for (int column = 0; column < columns; column++) widths[column] *= totalWidth / explicitTotal;
            explicitTotal = totalWidth;
        }

        double distributable = Math.Max(0D, totalWidth - explicitTotal - inferredColumns);
        double weightSum = 0D;
        for (int column = 0; column < columns; column++) {
            if (widths[column] == 0D) weightSum += weights[column];
        }
        for (int column = 0; column < columns; column++) {
            if (widths[column] == 0D) {
                widths[column] = 1D + (weightSum > 0D
                    ? distributable * weights[column] / weightSum
                    : distributable / inferredColumns);
            }
            table.ColumnWidths.Add(widths[column]);
        }
    }

    private static bool TryParseTableWidth(string? value, double referenceHalfInches, out double widthHalfInches) {
        widthHalfInches = 0D;
        if (TryParseCssPoints(value, out double points)) {
            widthHalfInches = points / 36D;
        } else {
            string text = (value ?? string.Empty).Trim();
            if (!text.EndsWith("%", StringComparison.Ordinal)
                || !double.TryParse(text.Substring(0, text.Length - 1), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double percent)) return false;
            widthHalfInches = referenceHalfInches * percent / 100D;
        }
        return widthHalfInches > 0D && !double.IsInfinity(widthHalfInches) && !double.IsNaN(widthHalfInches);
    }

    private static bool IsAuthoredWidth(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return !string.Equals(value!.Trim(), "auto", StringComparison.OrdinalIgnoreCase);
    }

    private static void ReportTableWidthApproximation(HtmlToOneNoteSectionResult result, string? authoredWidth) =>
        Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "An authored HTML table width could not be retained in OneNote's bounded page layout.",
            HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation, authoredWidth);

}
