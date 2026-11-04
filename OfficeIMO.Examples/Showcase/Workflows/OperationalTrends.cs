using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a six-week workload chart from editable worksheet values.</summary>
internal static class OperationalTrends {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        var sheet = document.AddSheet("Weekly trend");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(26)); sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(24)); sheet.SetColumnWidth(3, ExcelColumnWidth.Characters(24)); sheet.SetColumnWidth(4, ExcelColumnWidth.Characters(22));
        sheet.Range("A1:D1").Merge(); sheet.Cell(1, 1).SetValue("Support workload over six weeks");
        sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(20)); sheet.Cell(1, 1).SetBold(true); sheet.Cell(1, 1).SetFontColor("17365D");
        string[] headers = { "Week", "Received", "Resolved", "Net change" };
        for (int column = 1; column <= 4; column++) sheet.Cell(4, column).SetValue(headers[column - 1]);
        int[] received = { 82, 95, 88, 105, 97, 91 }, resolved = { 79, 90, 94, 98, 104, 100 };
        for (int index = 0; index < received.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1).SetValue("Week " + (index + 1)); sheet.Cell(row, 2).SetValue(received[index]); sheet.Cell(row, 3).SetValue(resolved[index]);
            sheet.Cell(row, 4).SetFormula($"B{row}-C{row}");
        }
        sheet.AddTable("A4:D10", hasHeader: true, name: "WeeklyWorkload", style: ExcelTableStyle.TableStyleMedium2);
        sheet.AddChart(OfficeChartType.Line, "A4:C10", new OfficeIMO.Excel.ExcelChartOptions { Row = 13, Column = 1, Width = OfficeLength.Pixels(620), Height = OfficeLength.Pixels(300), Title = "Received and resolved requests" });
        document.Calculate();
        sheet.Range("A1:D30").ExportImage(OfficeImageExportFormat.Png).Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
