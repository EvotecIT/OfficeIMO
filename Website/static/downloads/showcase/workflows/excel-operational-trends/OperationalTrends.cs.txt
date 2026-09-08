using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a six-week workload chart from editable worksheet values.</summary>
internal static class OperationalTrends {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        var sheet = document.AddWorksheet("Weekly trend");
        sheet.SetColumnWidth(1, 26); sheet.SetColumnWidth(2, 24); sheet.SetColumnWidth(3, 24); sheet.SetColumnWidth(4, 22);
        sheet.MergeRange("A1:D1"); sheet.Cell(1, 1, "Support workload over six weeks");
        sheet.CellFontSize(1, 1, 20); sheet.CellBold(1, 1, true); sheet.CellFontColor(1, 1, "17365D");
        string[] headers = { "Week", "Received", "Resolved", "Net change" };
        for (int column = 1; column <= 4; column++) sheet.Cell(4, column, headers[column - 1]);
        int[] received = { 82, 95, 88, 105, 97, 91 }, resolved = { 79, 90, 94, 98, 104, 100 };
        for (int index = 0; index < received.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1, "Week " + (index + 1)); sheet.Cell(row, 2, received[index]); sheet.Cell(row, 3, resolved[index]);
            sheet.CellFormula(row, 4, $"B{row}-C{row}");
        }
        sheet.AddTable("A4:D10", hasHeader: true, name: "WeeklyWorkload", style: ExcelTableStyle.TableStyleMedium2);
        sheet.AddChartFromRange("A4:C10", row: 13, column: 1, widthPixels: 620, heightPixels: 300,
            type: ExcelChartType.Line, title: "Received and resolved requests");
        document.Calculate();
        sheet.Range("A1:D30").ExportImage(OfficeImageExportFormat.Png).Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
