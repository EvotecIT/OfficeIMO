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
        FinishSheet(status, token);
        if (view.Rows.Any(r => r.BaselineStart.HasValue || r.BaselineFinish.HasValue)) {
            var baseline = document.AddWorksheet("Baseline dates");
            WriteHeaders(baseline, "UID", "Name", "Baseline start", "Baseline finish");
            for (int r = 0; r < view.Rows.Count; r++) {
                token.ThrowIfCancellationRequested(); var row = view.Rows[r];
                baseline.CellValue(r + 2, 1, row.Uid); baseline.CellValue(r + 2, 2, row.Name);
                if (row.BaselineStart.HasValue) SetExcelDate(baseline, r + 2, 3, row.BaselineStart.Value);
                if (row.BaselineFinish.HasValue) SetExcelDate(baseline, r + 2, 4, row.BaselineFinish.Value);
            }
            FinishSheet(baseline, token);
        }
        if (view.Links.Count > 0) {
            var dependencies = document.AddWorksheet("Dependencies");
            WriteHeaders(dependencies, "Predecessor UID", "Successor UID", "Type", "Lag");
            for (int r = 0; r < view.Links.Count; r++) {
                token.ThrowIfCancellationRequested(); var link = view.Links[r];
                dependencies.CellValue(r + 2, 1, link.PredecessorUid); dependencies.CellValue(r + 2, 2, link.SuccessorUid);
                dependencies.CellValue(r + 2, 3, link.Type.ToString()); dependencies.CellValue(r + 2, 4, link.LagText);
            }
            FinishSheet(dependencies, token);
        }
    }

    private static void WriteHeaders(ExcelSheet sheet, params string[] headers) {
        for (int c = 0; c < headers.Length; c++) sheet.CellValue(1, c + 1, headers[c]);
    }

    private static void SetExcelDate(ExcelSheet sheet, int row, int column, DateTime value) {
        sheet.CellValue(row, column, value); sheet.FormatCell(row, column, "yyyy-mm-dd hh:mm");
    }

    private static void FinishSheet(ExcelSheet sheet, CancellationToken token, int columns = 4) {
        sheet.Range("A1:" + sheet.CellAt(1, columns).Address).HeaderStyle();
        sheet.Freeze(1); sheet.AutoFitColumns(ct: token); sheet.ApplyPrintLayoutPreset(ExcelPrintLayoutPreset.Report);
    }
}
