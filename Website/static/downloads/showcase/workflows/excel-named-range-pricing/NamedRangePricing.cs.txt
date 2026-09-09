using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Uses named ranges to keep a small pricing model readable as the sheet evolves.</summary>
internal static class NamedRangePricing {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        ExcelSheet sheet = document.AddWorksheet("Quote");
        sheet.SetColumnWidth(1, 32);
        sheet.SetColumnWidth(2, 23);
        sheet.SetColumnWidth(3, 42);
        sheet.MergeRange("A1:C1");
        sheet.Cell(1, 1, "A quote with readable formulas");
        sheet.CellFontSize(1, 1, 20);
        sheet.CellBold(1, 1, true);
        sheet.CellFontColor(1, 1, "17365D");

        string[] labels = { "Quantity", "Unit price", "Discount rate", "Gross total", "Discount amount", "Net total" };
        string[] names = { "Quantity", "UnitPrice", "DiscountRate", "GrossTotal", "DiscountAmount", "NetTotal" };
        for (int index = 0; index < labels.Length; index++) {
            int row = index + 4;
            sheet.Cell(row, 1, labels[index]);
            sheet.Cell(row, 3, names[index]);
            sheet.CellFontColor(row, 3, "526179");
            document.SetNamedRange(names[index], $"'Quote'!B{row}", save: false);
        }
        sheet.Cell(4, 2, 24, numberFormat: "#,##0");
        sheet.Cell(5, 2, 125, numberFormat: "#,##0.00");
        sheet.Cell(6, 2, 0.10, numberFormat: "0%");
        sheet.CellFormula(7, 2, "Quantity*UnitPrice");
        sheet.CellFormula(8, 2, "GrossTotal*DiscountRate");
        sheet.CellFormula(9, 2, "GrossTotal-DiscountAmount");
        for (int row = 7; row <= 9; row++) {
            sheet.FormatCell(row, 2, "#,##0.00");
            sheet.CellBackground(row, 2, "E7F6ED");
        }
        sheet.CellBold(9, 2, true);
        sheet.MergeRange("A12:C12");
        sheet.Cell(12, 1, "Gross total, discount, and net total each have a named result.");
        document.Calculate();
        sheet.Range("A1:C14").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
