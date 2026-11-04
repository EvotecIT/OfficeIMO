using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Uses named ranges to keep a small pricing model readable as the sheet evolves.</summary>
internal static class NamedRangePricing {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        ExcelSheet sheet = document.AddSheet("Quote");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(32));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(23));
        sheet.SetColumnWidth(3, ExcelColumnWidth.Characters(42));
        sheet.Range("A1:C1").Merge();
        sheet.Cell(1, 1).SetValue("A quote with readable formulas");
        sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(20));
        sheet.Cell(1, 1).SetBold(true);
        sheet.Cell(1, 1).SetFontColor("17365D");

        string[] labels = { "Quantity", "Unit price", "Discount rate", "Gross total", "Discount amount", "Net total" };
        string[] names = { "Quantity", "UnitPrice", "DiscountRate", "GrossTotal", "DiscountAmount", "NetTotal" };
        for (int index = 0; index < labels.Length; index++) {
            int row = index + 4;
            sheet.Cell(row, 1).SetValue(labels[index]);
            sheet.Cell(row, 3).SetValue(names[index]);
            sheet.Cell(row, 3).SetFontColor("526179");
            document.SetNamedRange(names[index], $"'Quote'!B{row}", save: false);
        }
        sheet.Cell(4, 2).SetValue(24).SetNumberFormat("#,##0");
        sheet.Cell(5, 2).SetValue(125).SetNumberFormat("#,##0.00");
        sheet.Cell(6, 2).SetValue(0.10).SetNumberFormat("0%");
        sheet.Cell(7, 2).SetFormula("Quantity*UnitPrice");
        sheet.Cell(8, 2).SetFormula("GrossTotal*DiscountRate");
        sheet.Cell(9, 2).SetFormula("GrossTotal-DiscountAmount");
        for (int row = 7; row <= 9; row++) {
            sheet.Cell(row, 2).SetNumberFormat("#,##0.00");
            sheet.Cell(row, 2).SetFillColor("E7F6ED");
        }
        sheet.Cell(9, 2).SetBold(true);
        sheet.Range("A12:C12").Merge();
        sheet.Cell(12, 1).SetValue("Gross total, discount, and net total each have a named result.");
        document.Calculate();
        sheet.Range("A1:C14").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
