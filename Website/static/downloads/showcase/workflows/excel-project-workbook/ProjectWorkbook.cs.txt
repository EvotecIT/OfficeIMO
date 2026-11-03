using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Separates delivery detail from a summary with cross-sheet formulas.</summary>
internal static class ProjectWorkbook {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        ExcelSheet summary = document.AddSheet("Summary");
        ExcelSheet delivery = document.AddSheet("Delivery");
        summary.Range("A1:D1").Merge();
        summary.Cell(1, 1).SetValue("Project workbook / weekly delivery");
        summary.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(20));
        summary.Cell(1, 1).SetBold(true);
        summary.Cell(1, 1).SetFontColor("17365D");
        string[] labels = { "Discovery", "Build", "Pilot", "Handover" };
        string[] headers = { "Workstream", "Planned days", "Used days", "Remaining" };
        for (int column = 0; column < headers.Length; column++) {
            delivery.Cell(1, column + 1).SetValue(headers[column]);
            summary.Cell(4, column + 1).SetValue(headers[column]);
            summary.Cell(4, column + 1).SetFillColor("17365D");
            summary.Cell(4, column + 1).SetFontColor("FFFFFF");
            summary.SetColumnWidth(column + 1, ExcelColumnWidth.Characters(column == 0 ? 30 : 22));
            delivery.SetColumnWidth(column + 1, ExcelColumnWidth.Characters(column == 0 ? 30 : 22));
        }
        for (int index = 0; index < labels.Length; index++) {
            int row = index + 2;
            delivery.Cell(row, 1).SetValue(labels[index]);
            delivery.Cell(row, 2).SetValue(new[] { 12, 30, 10, 8 }[index]);
            delivery.Cell(row, 3).SetValue(new[] { 12, 18, 3, 0 }[index]);
            delivery.Cell(row, 4).SetFormula($"B{row}-C{row}");
            summary.Cell(index + 5, 1).SetValue(labels[index]);
            for (int column = 2; column <= 4; column++) {
                string letter = ((char)('A' + column - 1)).ToString();
                summary.Cell(index + 5, column).SetFormula($"'Delivery'!{letter}{row}");
            }
            summary.Cell(index + 5, 4).SetFillColor("E7F6ED");
        }
        delivery.AddTable("A1:D5", true, "DeliveryPlan", ExcelTableStyle.TableStyleMedium2);
        delivery.View.Freeze(topRows: 1);
        summary.Range("A11:D11").Merge();
        summary.Cell(11, 1).SetValue("Edit the Delivery sheet; Summary references the same values.");
        document.Calculate();
        summary.Range("A1:D13").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
