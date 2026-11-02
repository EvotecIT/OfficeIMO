using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates an editable asset depreciation register with visible calculations and a native range preview.</summary>
    internal static class AssetDepreciation {
        internal static void Create(string folder) {
            using var document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
            var sheet = document.AddSheet("Assets");
            // Widths sized to the content so the register prints on one page at full size.
            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(24));
            for (int column = 2; column <= 5; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(14));
            sheet.SetColumnWidth(6, ExcelColumnWidth.Characters(18));
            sheet.Range("A1:F1").Merge(); sheet.Cell(1, 1).SetValue("Asset depreciation register");
            sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(22)); sheet.Cell(1, 1).SetBold(true); sheet.Cell(1, 1).SetFontColor("17365D");
            sheet.Range("A2:F2").Merge(); sheet.Cell(2, 1).SetValue("Straight-line planning schedule / illustrative EUR values");
            string[] headers = { "Asset", "Cost", "Residual", "Life (years)", "Annual charge", "Book after year 1" };
            for (int column = 1; column <= 6; column++) sheet.Cell(4, column).SetValue(headers[column - 1]);
            string[] assets = { "Workshop laptops", "Meeting-room screens", "Storage equipment", "Lab instruments" };
            double[] costs = { 7200, 3600, 9200, 4800 }, residuals = { 600, 300, 800, 400 };
            int[] years = { 3, 4, 5, 4 };
            for (int index = 0; index < assets.Length; index++) {
                int row = index + 5;
                sheet.Cell(row, 1).SetValue(assets[index]); sheet.Cell(row, 2).SetValue(costs[index]); sheet.Cell(row, 3).SetValue(residuals[index]); sheet.Cell(row, 4).SetValue(years[index]);
                sheet.Cell(row, 5).SetFormula($"(B{row}-C{row})/D{row}");
                sheet.Cell(row, 6).SetFormula($"B{row}-E{row}");
                foreach (int column in new[] { 2, 3, 5, 6 }) sheet.Cell(row, column).SetNumberFormat("#,##0.00");
            }
            sheet.AddTable("A4:F8", true, "AssetSchedule", ExcelTableStyle.TableStyleMedium2);
            sheet.ValidationWholeNumber("D5:D8", ExcelDataValidationOperator.Between, 1, 50);
            sheet.Cell(10, 1).SetValue("Annual charge total"); sheet.Cell(10, 5).SetFormula("SUM(E5:E8)"); sheet.Cell(10, 5).SetNumberFormat("#,##0.00"); sheet.Cell(10, 5).SetBold(true);
            sheet.Range("A12:F12").Merge(); sheet.Cell(12, 1).SetValue("Assumes a full first year and cost above residual value.");
            sheet.Range("A13:F13").Merge(); sheet.Cell(13, 1).SetValue("Illustrative planning only; confirm accounting policy and tax treatment.");
            sheet.SetGridlinesVisible(false);
            for (int row = 1; row <= 16; row++) {
                sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(row == 1 ? 34 : row == 3 ? 12 : 25)));
                for (int column = 1; column <= 6; column++) {
                    if (row > 1) sheet.Cell(row, column).SetFontSize(OfficeLength.Points(11));
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
