using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates an editable cash-flow forecast with visible calculations and a native range preview.</summary>
    internal static class CashFlowForecast {
        internal static void Create(string folder) {
            using var document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
            var sheet = document.AddSheet("Cash flow");
            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(29));
            for (int column = 2; column <= 6; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(22));
            sheet.Range("A1:F1").Merge(); sheet.Cell(1, 1).SetValue("Cash-flow forecast");
            sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(22)); sheet.Cell(1, 1).SetBold(true); sheet.Cell(1, 1).SetFontColor("17365D");
            sheet.Range("A2:F2").Merge(); sheet.Cell(2, 1).SetValue("Six-month illustrative forecast / amounts in EUR");
            string[] headers = { "Month", "Opening", "Receipts", "Payments", "Net flow", "Closing" };
            for (int column = 1; column <= 6; column++) sheet.Cell(4, column).SetValue(headers[column - 1]);
            double[] receipts = { 24000, 26000, 22000, 29000, 27000, 31000 };
            double[] payments = { 21000, 28000, 25000, 24000, 26000, 27000 };
            for (int index = 0; index < 6; index++) {
                int row = index + 5;
                sheet.Cell(row, 1).SetValue(new[] { "January", "February", "March", "April", "May", "June" }[index]);
                if (index == 0) sheet.Cell(row, 2).SetValue(18000); else sheet.Cell(row, 2).SetFormula($"F{row - 1}");
                sheet.Cell(row, 3).SetValue(receipts[index]); sheet.Cell(row, 4).SetValue(payments[index]);
                sheet.Cell(row, 5).SetFormula($"C{row}-D{row}");
                sheet.Cell(row, 6).SetFormula($"B{row}+E{row}");
                for (int column = 2; column <= 6; column++) sheet.Cell(row, column).SetNumberFormat("#,##0");
            }
            sheet.AddTable("A4:F10", true, "CashForecast", ExcelTableStyle.TableStyleMedium2);
            sheet.ValidationWholeNumber("C5:D10", ExcelDataValidationOperator.Between, 0, 1000000);
            sheet.Range("A12:F12").Merge(); sheet.Cell(12, 1).SetValue("Edit receipts and payments; closing cash rolls into the following month.");
            sheet.Range("A13:F13").Merge(); sheet.Cell(13, 1).SetValue("Planning example only. Tax, financing and bank reconciliation are excluded.");
            sheet.SetGridlinesVisible(false);
            for (int row = 1; row <= 16; row++) {
                sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(row == 1 ? 34 : row == 3 ? 12 : 25)));
                for (int column = 1; column <= 6; column++) {
                    if (row > 1) sheet.Cell(row, column).SetFontSize(global::OfficeIMO.OfficeLength.Points(11));
                }
            }
            sheet.SetRowHeight(4, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
            for (int column = 1; column <= 6; column++) {
                sheet.Cell(4, column).SetFontColor("FFFFFF");
                sheet.Cell(4, column).SetBold(true);
                sheet.Cell(4, column).SetFillColor("17365D");
            }
            document.Calculate();
            document.InspectFormulas().EnsureAllSupported().EnsureAllHaveCachedResults().EnsureNoDependencyIssues();
            sheet.Range("A1:F14").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false })
                .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
            document.Save();
        }
    }
}
