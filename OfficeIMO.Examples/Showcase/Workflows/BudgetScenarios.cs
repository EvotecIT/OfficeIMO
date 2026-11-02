using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Compares three budget scenarios using calculated cells and editable assumptions.</summary>
internal static class BudgetScenarios {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        ExcelSheet sheet = document.AddSheet("Scenarios");
        sheet.Range("A1:F1").Merge();
        sheet.Cell(1, 1).SetValue("Budget scenarios / next quarter");
        sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(20));
        sheet.Cell(1, 1).SetBold(true);
        sheet.Cell(1, 1).SetFontColor("17365D");
        sheet.Range("A2:F2").Merge();
        sheet.Cell(2, 1).SetValue("Blue cells are assumptions. Green cells are calculated.");

        string[] headings = { "Scenario", "Units", "Unit price", "Revenue", "Fixed cost", "Contribution" };
        for (int column = 0; column < headings.Length; column++) {
            sheet.Cell(4, column + 1).SetValue(headings[column]);
            sheet.Cell(4, column + 1).SetFillColor("17365D");
            sheet.Cell(4, column + 1).SetFontColor("FFFFFF");
            sheet.Cell(4, column + 1).SetBold(true);
            sheet.SetColumnWidth(column + 1, ExcelColumnWidth.Characters(column == 0 ? 25 : 20));
        }
        string[] labels = { "Conservative", "Expected", "Growth" };
        for (int index = 0; index < labels.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1).SetValue(labels[index]);
            sheet.Cell(row, 2).SetValue(new[] { 800, 1100, 1500 }[index]).SetNumberFormat("#,##0");
            sheet.Cell(row, 3).SetValue(85).SetNumberFormat("#,##0.00");
            sheet.Cell(row, 4).SetFormula($"B{row}*C{row}");
            sheet.Cell(row, 5).SetValue(42000).SetNumberFormat("#,##0");
            sheet.Cell(row, 6).SetFormula($"D{row}-E{row}");
            sheet.Cell(row, 4).SetNumberFormat("#,##0");
            sheet.Cell(row, 6).SetNumberFormat("#,##0");
            foreach (int column in new[] { 2, 3, 5 }) sheet.Cell(row, column).SetFillColor("EAF1FB");
            foreach (int column in new[] { 4, 6 }) sheet.Cell(row, column).SetFillColor("E7F6ED");
        }
        sheet.Range("A10:F10").Merge();
        sheet.Cell(10, 1).SetValue("Change volume, price, or fixed cost to compare a different plan.");
        document.Calculate();
        sheet.Range("A1:F12").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
