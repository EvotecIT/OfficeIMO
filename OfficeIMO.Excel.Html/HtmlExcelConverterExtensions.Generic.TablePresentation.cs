using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void PreserveGenericTableCaption(
        ExcelSheet sheet,
        HtmlSemanticTable table,
        HtmlToExcelResult result,
        HtmlImportBudget budget) {
        A1.TryParseRange(sheet.UsedRangeA1, out _, out _, out int lastRow, out int lastColumn);
        foreach (ExcelMergedRangeSnapshot merge in sheet.GetMergedRanges(budget.Limits.MaxTableCells)) {
            lastRow = Math.Max(lastRow, merge.EndRow);
            lastColumn = Math.Max(lastColumn, merge.EndColumn);
        }
        if (lastRow < 1 || lastRow >= A1.MaxRows - 1) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                "The full table caption could not be placed after the native worksheet grid.",
                lossKind: OfficeConversionLossKind.Omission);
            return;
        }

        int captionRow = lastRow + 2;
        if (!TrySetCellTextValue(sheet, captionRow, 1, table.Caption, result, budget)) return;
        result.Cells++;
        if (lastColumn > 1) sheet.MergeRange(BuildCellReference(captionRow, 1) + ":" + BuildCellReference(captionRow, lastColumn));
        sheet.CellAt(captionRow, 1).SetBold();
        sheet.CellWrapText(captionRow, 1);
        sheet.AutoFitRow(captionRow);

        HtmlSemanticRun[] textRuns = table.CaptionRuns
            .Where(run => !string.IsNullOrWhiteSpace(run.Text))
            .ToArray();
        string[] linkTargets = textRuns
            .Select(run => run.Hyperlink)
            .Where(link => !string.IsNullOrWhiteSpace(link))
            .Select(link => link!)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (linkTargets.Length == 1 && textRuns.All(run => run.Hyperlink == linkTargets[0])) {
            if (budget.IsMetadataWithinLimit(linkTargets[0], out string hyperlinkLimit)) {
                sheet.SetHyperlinkReference(captionRow, 1, linkTargets[0], style: false);
            } else {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "The table caption hyperlink was omitted because its target exceeded the shared field limit.",
                    lossKind: OfficeConversionLossKind.Omission, detail: hyperlinkLimit);
            }
        } else if (linkTargets.Length > 0) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                "The table caption uses multiple or partially linked runs that one Excel cell cannot hyperlink faithfully.",
                lossKind: OfficeConversionLossKind.Omission,
                detail: "hyperlinkTargets=" + linkTargets.Length);
        }

        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "The authored table caption was placed after the native worksheet grid to keep its data coordinates.",
            lossKind: OfficeConversionLossKind.Approximation,
            detail: "cell=A" + captionRow + "; worksheet=" + sheet.Name
                + "; originalLength=" + table.Caption.Length);
    }

    private static void FormatGenericTableSheet(ExcelSheet sheet, HtmlSemanticTable table) {
        // The imported grid can contain empty source rows or stop at MaxTableCells.
        // Only format values that made it into the worksheet.
        ExcelCellValueInfo[] importedCells = sheet.EnumerateCells().ToArray();
        if (importedCells.Length == 0) return;

        if (table.Rows.Count > 0 && table.Rows.All(row => row.Cells.Count == 2
            && row.Cells.All(cell => cell.RowSpan == 1 && cell.ColumnSpan == 1))) {
            // Preserve the established printable term/definition presentation.
            FormatTwoColumnGenericTableSheet(sheet, importedCells);
        } else {
            // Use Excel's existing sizing rather than a separate HTML text
            // estimator. Bound long columns and fit rows after enabling wrapping.
            sheet.AutoFitColumnsFor(importedCells.Select(cell => cell.Column));
            ExcelColumnSnapshot[] columns = sheet.GetColumnDefinitions().ToArray();
            const double minimumWidth = 12D;
            double minimumTotal = columns.Sum(column => (column.EndIndex - column.StartIndex + 1) * minimumWidth);
            double extraTotal = columns.Sum(column => (column.EndIndex - column.StartIndex + 1)
                * (Math.Min(60D, Math.Max(minimumWidth, column.Width ?? minimumWidth)) - minimumWidth));
            // Reuse the existing printable width without squeezing a wide grid's
            // columns below the readable minimum. Only surplus width is shared.
            double extraBudget = Math.Max(75D, minimumTotal) - minimumTotal;
            double scale = extraTotal > 0D ? Math.Min(1D, extraBudget / extraTotal) : 1D;
            var fittedWidths = new Dictionary<int, double>();
            foreach (ExcelColumnSnapshot column in columns) {
                double width = minimumWidth
                    + (Math.Min(60D, Math.Max(minimumWidth, column.Width ?? minimumWidth)) - minimumWidth) * scale;
                for (int index = column.StartIndex; index <= column.EndIndex; index++) {
                    fittedWidths.Add(index, width);
                }
            }
            sheet.SetColumnWidths(fittedWidths);
        }

        sheet.CellWrapTextFor(importedCells.Select(cell => (cell.Row, cell.Column)));
        sheet.AutoFitRows();
    }

    private static void FormatTwoColumnGenericTableSheet(ExcelSheet sheet, ExcelCellValueInfo[] importedCells) {
        int firstLength = 0;
        int secondLength = 0;
        foreach (ExcelCellValueInfo cell in importedCells) {
            int length = cell.Value?.ToString()?.Length ?? 0;
            if (cell.Column == 1) firstLength = Math.Max(firstLength, length);
            else secondLength = Math.Max(secondLength, length);
        }

        double firstWidth = Math.Min(60D, Math.Max(12D, firstLength + 2D));
        double secondWidth = Math.Min(60D, Math.Max(12D, secondLength + 2D));
        const double printableWidth = 75D;
        if (firstWidth + secondWidth > printableWidth) {
            firstWidth = Math.Max(12D, Math.Ceiling(firstWidth * printableWidth / (firstWidth + secondWidth)));
            secondWidth = printableWidth - firstWidth;
        }

        sheet.SetColumnWidth(1, firstWidth);
        sheet.SetColumnWidth(2, secondWidth);
    }
}
