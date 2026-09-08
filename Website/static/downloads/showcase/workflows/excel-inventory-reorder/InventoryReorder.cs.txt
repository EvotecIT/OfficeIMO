using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Calculates a replenishment proposal from on-hand stock, targets and unit costs.</summary>
internal static class InventoryReorder {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        var sheet = document.AddWorksheet("Reorder");
        sheet.SetColumnWidth(1, 30);
        for (int column = 2; column <= 6; column++) sheet.SetColumnWidth(column, 15);
        sheet.SetColumnWidth(7, 19);
        sheet.MergeRange("A1:G1");
        sheet.Cell(1, 1, "Inventory replenishment plan");
        sheet.CellFontSize(1, 1, 20); sheet.CellBold(1, 1, true); sheet.CellFontColor(1, 1, "17365D");
        sheet.MergeRange("A2:G2"); sheet.Cell(2, 1, "Illustrative stock snapshot / quantities and unit costs are editable");
        string[] headers = { "Item", "On hand", "Target", "Stock gap", "Order qty", "Unit cost", "Order cost" };
        for (int column = 1; column <= 7; column++) sheet.Cell(4, column, headers[column - 1]);
        string[] items = { "USB-C dock", "Wireless keyboard", "Headset", "Laptop stand", "Travel adapter" };
        int[] onHand = { 8, 24, 5, 12, 3 }, targets = { 20, 20, 15, 12, 10 };
        double[] costs = { 115, 42, 64, 28, 19 };
        for (int index = 0; index < items.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1, items[index]); sheet.Cell(row, 2, onHand[index]); sheet.Cell(row, 3, targets[index]);
            sheet.CellFormula(row, 4, $"C{row}-B{row}");
            sheet.CellFormula(row, 5, $"IF(D{row}>0,D{row},0)");
            sheet.Cell(row, 6, costs[index], numberFormat: "#,##0.00");
            sheet.CellFormula(row, 7, $"E{row}*F{row}"); sheet.FormatCell(row, 7, "#,##0.00");
        }
        sheet.AddTable("A4:G9", hasHeader: true, name: "Stock", style: ExcelTableStyle.TableStyleMedium2);
        sheet.ValidationWholeNumber("B5:C9", ExcelDataValidationOperator.Between, 0, 10000);
        sheet.Cell(11, 1, "Proposed spend"); sheet.CellFormula(11, 7, "SUM(G5:G9)"); sheet.FormatCell(11, 7, "#,##0.00");
        sheet.CellBold(11, 7, true); sheet.CellBackground(11, 7, "E7F6ED");
        sheet.MergeRange("A13:G13"); sheet.Cell(13, 1, "Proposal only: check lead times and committed orders before purchasing.");
        document.Calculate();
        sheet.Range("A1:G15").ExportImage(OfficeImageExportFormat.Png).Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
