using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates an editable weekly support rota with visible calculations and a native range preview.</summary>
    internal static class WorkforceRota {
        internal static void Create(string folder) {
            using var document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
            var sheet = document.AddSheet("Rota");
            // Day columns hold one letter; narrow widths keep the rota on one page at full size.
            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(18));
            for (int column = 2; column <= 6; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(9));
            sheet.SetColumnWidth(7, ExcelColumnWidth.Characters(12));
            sheet.Range("A1:G1").Merge(); sheet.Cell(1, 1).SetValue("Weekly support rota");
            sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(22)); sheet.Cell(1, 1).SetBold(true); sheet.Cell(1, 1).SetFontColor("17365D");
            sheet.Range("A2:G2").Merge(); sheet.Cell(2, 1).SetValue("One assigned shift per day / D = desk / R = remote / - = off");
            string[] headers = { "Team member", "Mon", "Tue", "Wed", "Thu", "Fri", "Desk days" };
            for (int column = 1; column <= 7; column++) sheet.Cell(4, column).SetValue(headers[column - 1]);
            string[] names = { "Maya", "Leon", "Sofia", "Amir" };
            string[][] shifts = { new[] { "D", "D", "R", "-", "R" }, new[] { "R", "-", "D", "D", "R" }, new[] { "D", "R", "-", "R", "D" }, new[] { "-", "D", "R", "D", "D" } };
            for (int index = 0; index < names.Length; index++) {
                int row = index + 5; sheet.Cell(row, 1).SetValue(names[index]);
                for (int column = 2; column <= 6; column++) {
                    string shift = shifts[index][column - 2]; sheet.Cell(row, column).SetValue(shift);
                }
                sheet.Cell(row, 7).SetFormula($"COUNTIF(B{row}:F{row},\"D\")");
            }
            sheet.AddTable("A4:G8", true, "SupportRota", ExcelTableStyle.TableStyleMedium2);
            // Rules follow edited shifts in Excel and in the native range preview.
            sheet.Range("B5:F8").Validation().List(new[] { "D", "R", "-" }, allowBlank: false);
            sheet.Range("B5:F8").ConditionalFormat().Text(ExcelConditionalFormatType.ContainsText, "D", "DCFCE7");
            sheet.Range("B5:F8").ConditionalFormat().Text(ExcelConditionalFormatType.ContainsText, "R", "DBEAFE");
            sheet.Range("B5:F8").ConditionalFormat().Text(ExcelConditionalFormatType.ContainsText, "-", "F1F5F9");
            sheet.Cell(10, 1).SetValue("Desk coverage");
            for (int column = 2; column <= 6; column++) {
                string letter = ((char)('A' + column - 1)).ToString();
                sheet.Cell(10, column).SetFormula($"COUNTIF({letter}5:{letter}8,\"D\")"); sheet.Cell(10, column).SetBold(true);
            }
            sheet.Range("A12:G12").Merge(); sheet.Cell(12, 1).SetValue("Review coverage before sharing. Names and shifts are synthetic.");
            sheet.Range("A13:G13").Merge(); sheet.Cell(13, 1).SetValue("This records a rota; it does not check working-time rules or approve leave.");
            sheet.View.ShowGridlines = false;
            for (int row = 1; row <= 16; row++) {
                sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(row == 1 ? 34 : row == 3 ? 12 : 25)));
                for (int column = 1; column <= 7; column++) {
                    if (row > 1) sheet.Cell(row, column).SetFontSize(OfficeLength.Points(11));
                }
            }
            sheet.SetRowHeight(4, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
            for (int column = 1; column <= 7; column++) {
                sheet.Cell(4, column).SetFontColor("FFFFFF");
                sheet.Cell(4, column).SetBold(true);
                sheet.Cell(4, column).SetFillColor("17365D");
            }
            document.Calculate();
            document.InspectFormulas().EnsureAllSupported().EnsureAllHaveCachedResults().EnsureNoDependencyIssues();
            sheet.Range("A1:G14").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false })
                .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
            document.Save();
        }
    }
}
