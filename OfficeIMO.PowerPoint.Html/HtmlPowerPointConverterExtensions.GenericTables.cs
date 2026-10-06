using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static bool TryCreatePagedGenericTableGrid(
        HtmlSemanticBlock block,
        IReadOnlyList<double> rowHeights,
        double slideBottom,
        HtmlImportBudget budget,
        HtmlToPowerPointResult result,
        out PowerPointHtmlTableGrid? grid) {
        grid = null;
        HtmlSemanticTable? source = block.Table;
        if (source == null || source.Rows.Count < 2
            || block.SourceElement.HasAttribute("data-officeimo-left")
            || block.SourceElement.HasAttribute("data-officeimo-top")
            || block.SourceElement.HasAttribute("data-officeimo-height")
            || source.Rows.Any(row => row.Cells.Any(cell => cell.RowSpan != 1))) return false;

        double available = slideBottom - 30D;
        bool repeatHeader = HasGenericTableHeaderRow(source);
        for (int row = 0; row < rowHeights.Count; row++) {
            double required = rowHeights[row] + (repeatHeader && row > 0 ? rowHeights[0] : 0D);
            if (double.IsNaN(required) || double.IsInfinity(required)
                || Math.Max(90D, required) + 20D > available) return false;
        }

        PowerPointHtmlTableGrid candidate = BuildTableGrid(block.SourceElement, budget, result);
        if (candidate.Rows != rowHeights.Count || candidate.Columns == 0
            || candidate.Cells.Any(cell => cell.RowSpan != 1)) return false;
        grid = candidate;
        return true;
    }

    private static bool HasGenericTableHeaderRow(HtmlSemanticTable source) =>
        source.Rows.Count > 1 && source.Rows[0].SourceRowIndex == 0 && source.Rows[0].Cells.Count > 0
        && source.Rows[0].Cells.All(cell => cell.IsHeader);

    private static bool TryImportPagedGenericTable(
        HtmlSemanticBlock block,
        PowerPointHtmlTableGrid grid,
        IReadOnlyList<double> rowHeights,
        double tableWidth,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        ref PptCore.PowerPointSlide slide,
        ref double contentTop,
        ref double pictureTop,
        double slideBottom) {
        HtmlSemanticTable source = block.Table!;
        bool repeatHeader = HasGenericTableHeaderRow(source);
        int nextRow = 0;
        int segments = 0;
        bool reservedTable = false;
        while (nextRow < grid.Rows) {
            bool includeHeader = repeatHeader && nextRow > 0;
            double available = slideBottom - contentTop - 20D;
            double usedHeight = includeHeader ? rowHeights[0] : 0D;
            int endRow = nextRow;
            while (endRow < grid.Rows
                   && Math.Max(90D, usedHeight + rowHeights[endRow]) <= available) {
                usedHeight += rowHeights[endRow];
                endRow++;
            }
            // A repeated-header table must never strand its first header without a data row.
            if (endRow == nextRow || repeatHeader && nextRow == 0 && endRow == 1) {
                if (!reservedTable && !TryReservePagedGenericTable(budget, result, nextRow)) return true;
                reservedTable = true;
                if (!TryAddGenericSlide(presentation, result, budget, out slide)) return false;
                contentTop = pictureTop = 30D;
                continue;
            }
            if (!reservedTable && !TryReservePagedGenericTable(budget, result, nextRow)) return true;
            reservedTable = false;

            int rows = endRow - nextRow + (includeHeader ? 1 : 0);
            double height = Math.Max(90D, usedHeight);
            PptCore.PowerPointTable table = slide.AddTablePoints(rows, grid.Columns,
                64D, contentTop, tableWidth, height);
            int localRow = 0;
            if (includeHeader) {
                CopyGenericTableRow(table, grid, source, 0, localRow++, options, result);
                table.SetRowHeightPoints(0, rowHeights[0]);
            }
            for (int sourceRow = nextRow; sourceRow < endRow; sourceRow++, localRow++) {
                CopyGenericTableRow(table, grid, source, sourceRow, localRow, options, result);
                double padding = localRow == rows - 1 ? height - usedHeight : 0D;
                table.SetRowHeightPoints(localRow, rowHeights[sourceRow] + padding);
            }
            ApplyShapeTransforms(block.SourceElement, table, budget, result);
            result.Tables++;
            segments++;
            contentTop += height + 20D;
            pictureTop = Math.Max(pictureTop, contentTop);
            nextRow = endRow;
        }
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "An HTML table was partitioned into editable row-group tables across slides.",
            lossKind: OfficeConversionLossKind.Approximation,
            detail: "projection=paginatedNativeTable; segments=" + segments
                + "; repeatedHeader=" + repeatHeader);
        return true;
    }

    private static bool TryReservePagedGenericTable(
        HtmlImportBudget budget,
        HtmlToPowerPointResult result,
        int nextRow) {
        if (budget.TryReserveTableWithShape(out string tableLimit)) return true;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
            "Remaining HTML table rows were omitted because the shared table or shape limit was reached.",
            lossKind: OfficeConversionLossKind.Omission,
            detail: tableLimit + "; firstOmittedRow=" + (nextRow + 1));
        return false;
    }

    private static void CopyGenericTableRow(
        PptCore.PowerPointTable table,
        PowerPointHtmlTableGrid grid,
        HtmlSemanticTable source,
        int sourceRow,
        int targetRow,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result) {
        HtmlSemanticTableRow? semanticRow = source.Rows.FirstOrDefault(row => row.SourceRowIndex == sourceRow);
        int sourceCell = 0;
        foreach (PowerPointHtmlTableCell cell in grid.Cells.Where(cell => cell.Row == sourceRow)) {
            PptCore.PowerPointTableCell targetCell = table.GetCell(targetRow, cell.Column);
            targetCell.Text = cell.Text;
            if (cell.ColumnSpan > 1) {
                table.MergeCells(targetRow, cell.Column, targetRow, cell.Column + cell.ColumnSpan - 1);
                result.MergedRanges++;
            }
            if (semanticRow != null && sourceCell < semanticRow.Cells.Count) {
                ApplySemanticTableCellFormatting(targetCell, semanticRow.Cells[sourceCell],
                    options.NormalizedHyperlinkUrlPolicy ?? options.HyperlinkUrlPolicy,
                    table.FirstRow && targetRow == 0);
            }
            TryApplyTargetSemanticRuns(targetCell, cell.Element, options.HyperlinkUrlPolicy);
            sourceCell++;
        }
    }
}
