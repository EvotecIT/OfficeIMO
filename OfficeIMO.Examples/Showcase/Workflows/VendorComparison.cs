using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates an editable vendor evaluation scorecard with visible calculations and a native range preview.</summary>
    internal static class VendorComparison {
        internal static void Create(string folder) {
            using var document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
            var sheet = document.AddSheet("Evaluation");
            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(29));
            for (int column = 2; column <= 5; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(22));
            sheet.Range("A1:E1").Merge(); sheet.Cell(1, 1).SetValue("Vendor evaluation scorecard");
            sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(22)); sheet.Cell(1, 1).SetBold(true); sheet.Cell(1, 1).SetFontColor("17365D");
            sheet.Range("A2:E2").Merge(); sheet.Cell(2, 1).SetValue("Weighted scoring / higher scores are better / synthetic offers");
            string[] headers = { "Vendor", "Quality (50%)", "Delivery (30%)", "Support (20%)", "Weighted score" };
            for (int column = 1; column <= 5; column++) sheet.Cell(4, column).SetValue(headers[column - 1]);
            string[] vendors = { "Cedar Services", "Harbour Systems", "Juniper Partners" };
            int[,] scores = { { 4, 3, 5 }, { 3, 5, 4 }, { 5, 4, 3 } };
            for (int index = 0; index < vendors.Length; index++) {
                int row = index + 5; sheet.Cell(row, 1).SetValue(vendors[index]);
                for (int column = 2; column <= 4; column++) sheet.Cell(row, column).SetValue(scores[index, column - 2]);
                sheet.Cell(row, 5).SetFormula($"B{row}*0.5+C{row}*0.3+D{row}*0.2"); sheet.Cell(row, 5).SetNumberFormat("0.00");
            }
            sheet.AddTable("A4:E7", true, "VendorScores", ExcelTableStyle.TableStyleMedium2);
            sheet.ValidationWholeNumber("B5:D7", ExcelDataValidationOperator.Between, 1, 5);
            sheet.Range("A10:E10").Merge(); sheet.Cell(10, 1).SetValue("Scoring guide: 1 = limited evidence; 3 = meets need; 5 = strong evidence.");
            sheet.Range("A12:E12").Merge(); sheet.Cell(12, 1).SetValue("Weights sum to 100%. Keep evidence and conflicts of interest in the review record.");
            sheet.Range("A13:E13").Merge(); sheet.Cell(13, 1).SetValue("Scores support review; they do not approve a purchase.");
            sheet.SetGridlinesVisible(false);
            for (int row = 1; row <= 16; row++) {
                sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(row == 1 ? 34 : row == 3 ? 12 : 25)));
                for (int column = 1; column <= 5; column++) {
                    if (row > 1) sheet.Cell(row, column).SetFontSize(global::OfficeIMO.OfficeLength.Points(11));
                }
            }
            sheet.SetRowHeight(4, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
            for (int column = 1; column <= 5; column++) {
                sheet.Cell(4, column).SetFontColor("FFFFFF");
                sheet.Cell(4, column).SetBold(true);
                sheet.Cell(4, column).SetFillColor("17365D");
            }
            document.Calculate();
            document.InspectFormulas().EnsureAllSupported().EnsureAllHaveCachedResults().EnsureNoDependencyIssues();
            sheet.Range("A1:E14").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false })
                .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
            document.Save();
        }
    }
}
