using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void PreserveGenericTableCaption(
        ExcelSheet sheet,
        HtmlSemanticTable table,
        HtmlToExcelResult result,
        HtmlImportBudget budget) {
        if (!A1.TryParseRange(sheet.UsedRangeA1, out _, out _, out int lastRow, out _)
            || lastRow >= A1.MaxRows - 1) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                "The full table caption could not be placed after the native worksheet grid.",
                lossKind: OfficeConversionLossKind.Omission);
            return;
        }

        int captionRow = lastRow + 2;
        if (!TrySetCellTextValue(sheet, captionRow, 1, table.Caption, result, budget)) return;
        result.Cells++;
        sheet.CellAt(captionRow, 1).SetBold();
        sheet.CellWrapText(captionRow, 1);

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
            sheet.SetHyperlinkReference(captionRow, 1, linkTargets[0], style: false);
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

    private static void FormatSimpleGenericTableSheet(ExcelSheet sheet, HtmlSemanticTable table) {
        // Generic two-column tables are often term/definition data. Keep both
        // columns within a printable width and let longer values wrap visibly.
        if (table.Rows.Count == 0 || table.Rows.Any(row => row.Cells.Count != 2
            || row.Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1))) return;
        // The imported grid can contain empty source rows or stop at MaxTableCells.
        // Only format values that made it into the worksheet.
        ExcelCellValueInfo[] importedCells = sheet.EnumerateCells()
            .Where(cell => cell.Column is 1 or 2)
            .ToArray();
        if (importedCells.Length == 0) return;

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
        foreach (ExcelCellValueInfo cell in importedCells) {
            sheet.CellWrapText(cell.Row, cell.Column);
        }
    }
}
