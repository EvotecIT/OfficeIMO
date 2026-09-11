using OfficeIMO.Excel;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    private static void AddExcelSupplementalTables(ExcelDocument document, ProjectView view, CancellationToken token) {
        var status = document.AddWorksheet("Status and groups");
        WriteHeaders(status, "UID", "Name", "Status", "Group");
        for (int r = 0; r < view.Rows.Count; r++) {
            token.ThrowIfCancellationRequested(); var row = view.Rows[r];
            status.CellValue(r + 2, 1, row.Uid); status.CellValue(r + 2, 2, row.Name);
            status.CellValue(r + 2, 3, Status(view, row)); status.CellValue(r + 2, 4, row.Group);
        }
        FinishSheet(status, token, view.Rows.Count + 1);
        if (view.Rows.Any(r => r.BaselineStart.HasValue || r.BaselineFinish.HasValue)) {
            var baseline = document.AddWorksheet("Baseline dates");
            WriteHeaders(baseline, "UID", "Name", "Baseline start", "Baseline finish");
            for (int r = 0; r < view.Rows.Count; r++) {
                token.ThrowIfCancellationRequested(); var row = view.Rows[r];
                baseline.CellValue(r + 2, 1, row.Uid); baseline.CellValue(r + 2, 2, row.Name);
                if (row.BaselineStart.HasValue) SetExcelDate(baseline, r + 2, 3, row.BaselineStart.Value);
                if (row.BaselineFinish.HasValue) SetExcelDate(baseline, r + 2, 4, row.BaselineFinish.Value);
            }
            FinishSheet(baseline, token, view.Rows.Count + 1);
        }
        if (view.Links.Count > 0) {
            var dependencies = document.AddWorksheet("Dependencies");
            WriteHeaders(dependencies, "Predecessor UID", "Successor UID", "Type", "Lag");
            for (int r = 0; r < view.Links.Count; r++) {
                token.ThrowIfCancellationRequested(); var link = view.Links[r];
                dependencies.CellValue(r + 2, 1, link.PredecessorUid); dependencies.CellValue(r + 2, 2, link.SuccessorUid);
                dependencies.CellValue(r + 2, 3, DependencyText(link.Type)); dependencies.CellValue(r + 2, 4, link.LagText);
            }
            FinishSheet(dependencies, token, view.Links.Count + 1);
        }
    }

    private static void WriteHeaders(ExcelSheet sheet, params string[] headers) {
        for (int c = 0; c < headers.Length; c++) sheet.CellValue(1, c + 1, headers[c]);
    }

    private static void SetExcelDate(ExcelSheet sheet, int row, int column, DateTime value) {
        sheet.CellValue(row, column, value); sheet.FormatCell(row, column, "yyyy-mm-dd hh:mm");
    }

    private static void FinishSheet(ExcelSheet sheet, CancellationToken token, int rows, int columns = 4, int repeatedColumns = 0) {
        sheet.Range(sheet.UsedRangeA1).SetFontName("Arial").SetFontSize(12).SetFontColor("183047");
        sheet.Range("A1:" + sheet.CellAt(1, columns).Address).SetBold().SetFillColor("183047").SetFontColor("FFFFFF");
        sheet.Freeze(1); sheet.AutoFitColumns(ct: token);
        for (int c = 1; c <= columns; c++) {
            token.ThrowIfCancellationRequested();
            sheet.TryGetCellText(1, c, out string? header);
            if (header == "Name" || header == "Group") sheet.WrapCells(1, rows, c, header == "Name" ? 40 : 24);
            else if (header == "Cost" || header == "Critical") sheet.SetColumnWidth(c, 12);
            else if (repeatedColumns > 0 && c > repeatedColumns) sheet.SetColumnWidth(c, 14);
        }
        sheet.SetRowLayout(1, new ExcelRowLayoutOptions { WrapText = true, FirstColumn = 1, LastColumn = columns });
        if (rows > 1) sheet.SetRowsLayout(Enumerable.Range(2, rows - 1).Where(row => row % 2 == 0),
            new ExcelRowLayoutOptions { BackgroundColor = "F4F7FA", FirstColumn = 1, LastColumn = columns });
        sheet.AutoFitRows(ct: token);
        sheet.ApplyPrintLayout(new ExcelPrintLayoutOptions {
            Preset = columns > 8 ? ExcelPrintLayoutPreset.Worksheet : ExcelPrintLayoutPreset.Report, PaperSize = ExcelPaperSize.A4,
            Orientation = columns <= 5 ? OfficePageOrientation.Portrait : OfficePageOrientation.Landscape,
            Margins = ExcelMarginPreset.Narrow, RepeatFirstRow = 1, RepeatLastRow = 1,
            RepeatFirstColumn = repeatedColumns > 0 ? 1 : null, RepeatLastColumn = repeatedColumns > 0 ? repeatedColumns : null
        });
        sheet.SetPrintOptions(horizontalCentered: true);
    }
}
