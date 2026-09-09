using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Compares three budget scenarios using calculated cells and editable assumptions.</summary>
internal static class BudgetScenarios {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        ExcelSheet sheet = document.AddWorksheet("Scenarios");
        sheet.MergeRange("A1:F1");
        sheet.Cell(1, 1, "Budget scenarios / next quarter");
        sheet.CellFontSize(1, 1, 20);
        sheet.CellBold(1, 1, true);
        sheet.CellFontColor(1, 1, "17365D");
        sheet.MergeRange("A2:F2");
        sheet.Cell(2, 1, "Blue cells are assumptions. Green cells are calculated.");

        string[] headings = { "Scenario", "Units", "Unit price", "Revenue", "Fixed cost", "Contribution" };
        for (int column = 0; column < headings.Length; column++) {
            sheet.Cell(4, column + 1, headings[column]);
            sheet.CellBackground(4, column + 1, "17365D");
            sheet.CellFontColor(4, column + 1, "FFFFFF");
            sheet.CellBold(4, column + 1, true);
            sheet.SetColumnWidth(column + 1, column == 0 ? 25 : 20);
        }
        string[] labels = { "Conservative", "Expected", "Growth" };
        for (int index = 0; index < labels.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1, labels[index]);
            sheet.Cell(row, 2, new[] { 800, 1100, 1500 }[index], numberFormat: "#,##0");
            sheet.Cell(row, 3, 85, numberFormat: "#,##0.00");
            sheet.CellFormula(row, 4, $"B{row}*C{row}");
            sheet.Cell(row, 5, 42000, numberFormat: "#,##0");
            sheet.CellFormula(row, 6, $"D{row}-E{row}");
            sheet.FormatCell(row, 4, "#,##0");
            sheet.FormatCell(row, 6, "#,##0");
            foreach (int column in new[] { 2, 3, 5 }) sheet.CellBackground(row, column, "EAF1FB");
            foreach (int column in new[] { 4, 6 }) sheet.CellBackground(row, column, "E7F6ED");
        }
        sheet.MergeRange("A10:F10");
        sheet.Cell(10, 1, "Change volume, price, or fixed cost to compare a different plan.");
        document.Calculate();
        sheet.Range("A1:F12").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
