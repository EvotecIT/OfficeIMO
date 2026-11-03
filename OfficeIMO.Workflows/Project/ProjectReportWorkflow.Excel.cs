using OfficeIMO.Excel;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    private static void ValidateExcelGrid(ProjectView view) {
        if (view.Rows.Count > ExcelA1.MaxRows - 1)
            throw new InvalidOperationException("The report has more rows than an Excel worksheet can contain.");
        if (view.Links.Count > ExcelA1.MaxRows - 1)
            throw new InvalidOperationException("The dependency table has more rows than an Excel worksheet can contain.");
        if (IsUsage(view) && view.Buckets.Count > ExcelA1.MaxColumns - 2)
            throw new InvalidOperationException("The usage table has more time buckets than an Excel worksheet can contain.");
    }

    private static void AddExcelSupplementalTables(ExcelDocument document, ProjectView view, CancellationToken token) {
        var status = document.AddSheet("Status and groups");
        WriteHeaders(status, "UID", "Name", "Status", "Group");
        for (int r = 0; r < view.Rows.Count; r++) {
            token.ThrowIfCancellationRequested(); var row = view.Rows[r];
            status.Cell(r + 2, 1).SetValue(row.Uid); status.Cell(r + 2, 2).SetValue(row.Name);
            status.Cell(r + 2, 3).SetValue(Status(view, row)); status.Cell(r + 2, 4).SetValue(row.Group);
        }
        FinishSheet(status, token, view.Rows.Count + 1);
        if (view.Rows.Any(r => r.BaselineStart.HasValue || r.BaselineFinish.HasValue)) {
            var baseline = document.AddSheet("Baseline dates");
            WriteHeaders(baseline, "UID", "Name", "Baseline start", "Baseline finish");
            for (int r = 0; r < view.Rows.Count; r++) {
                token.ThrowIfCancellationRequested(); var row = view.Rows[r];
                baseline.Cell(r + 2, 1).SetValue(row.Uid); baseline.Cell(r + 2, 2).SetValue(row.Name);
                if (row.BaselineStart.HasValue) SetExcelDate(baseline, r + 2, 3, row.BaselineStart.Value);
                if (row.BaselineFinish.HasValue) SetExcelDate(baseline, r + 2, 4, row.BaselineFinish.Value);
            }
            FinishSheet(baseline, token, view.Rows.Count + 1);
        }
        if (view.Links.Count > 0) {
            var dependencies = document.AddSheet("Dependencies");
            WriteHeaders(dependencies, "Predecessor UID", "Successor UID", "Type", "Lag");
            for (int r = 0; r < view.Links.Count; r++) {
                token.ThrowIfCancellationRequested(); var link = view.Links[r];
                dependencies.Cell(r + 2, 1).SetValue(link.PredecessorUid); dependencies.Cell(r + 2, 2).SetValue(link.SuccessorUid);
                dependencies.Cell(r + 2, 3).SetValue(DependencyText(link.Type)); dependencies.Cell(r + 2, 4).SetValue(link.LagText);
            }
            FinishSheet(dependencies, token, view.Links.Count + 1);
        }
    }

    private static void WriteHeaders(ExcelSheet sheet, params string[] headers) {
        for (int c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).SetValue(headers[c]);
    }

    private static void SetExcelDate(ExcelSheet sheet, int row, int column, DateTime value) {
        sheet.Cell(row, column).SetValue(value); sheet.Cell(row, column).SetNumberFormat("yyyy-mm-dd hh:mm");
    }

    private static void FinishSheet(ExcelSheet sheet, CancellationToken token, int rows, int columns = 4, int repeatedColumns = 0) {
        sheet.UsedRange.SetFontName("Arial").SetFontSize(global::OfficeIMO.OfficeLength.Points(12)).SetFontColor("183047");
        sheet.Range("A1:" + sheet.Cell(1, columns).Address).SetBold().SetFillColor("183047").SetFontColor("FFFFFF");
        sheet.View.Freeze(1); sheet.AutoFitColumns(ct: token);
        for (int c = 1; c <= columns; c++) {
            token.ThrowIfCancellationRequested();
            sheet.TryGetCellText(1, c, out string? header);
            if (header == "Name" || header == "Group") {
                sheet.Range(1, c, rows, c).SetWrapText();
                sheet.SetColumnWidth(c, ExcelColumnWidth.Characters(header == "Name" ? 40 : 24));
            }
            else if (header == "Cost" || header == "Critical") sheet.SetColumnWidth(c, ExcelColumnWidth.Characters(12));
            else if (repeatedColumns > 0 && c > repeatedColumns) sheet.SetColumnWidth(c, ExcelColumnWidth.Characters(14));
        }
        sheet.SetRowLayout(1, new ExcelRowLayoutOptions { WrapText = true, FirstColumn = 1, LastColumn = columns });
        if (rows > 1) sheet.SetRowsLayout(Enumerable.Range(2, rows - 1).Where(row => row % 2 == 0),
            new ExcelRowLayoutOptions { BackgroundColor = "F4F7FA", FirstColumn = 1, LastColumn = columns });
        sheet.AutoFitRows(ct: token);
        sheet.PageSetup.ApplyLayout(new ExcelPrintLayoutOptions {
            Preset = columns > 8 ? ExcelPrintLayoutPreset.Worksheet : ExcelPrintLayoutPreset.Report, PaperSize = ExcelPaperSize.A4,
            Orientation = columns <= 5 ? OfficePageOrientation.Portrait : OfficePageOrientation.Landscape,
            Margins = ExcelMarginPreset.Narrow, RepeatFirstRow = 1, RepeatLastRow = 1,
            RepeatFirstColumn = repeatedColumns > 0 ? 1 : null, RepeatLastColumn = repeatedColumns > 0 ? repeatedColumns : null
        });
        sheet.PageSetup.HorizontalCentered = true;
    }
}
