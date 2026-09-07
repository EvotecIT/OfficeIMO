using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds an expense register with real date cells, currency formats, and list validation.</summary>
internal static class ExpenseRegister {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        ExcelSheet sheet = document.AddWorksheet("Expenses");
        sheet.MergeRange("A1:E1");
        sheet.Cell(1, 1, "September expense register");
        sheet.CellFontSize(1, 1, 20);
        sheet.CellBold(1, 1, true);
        sheet.CellFontColor(1, 1, "17365D");
        string[] headers = { "Date", "Description", "Category", "Amount (EUR)", "Status" };
        for (int column = 0; column < headers.Length; column++) {
            sheet.Cell(4, column + 1, headers[column]);
            sheet.SetColumnWidth(column + 1, column == 1 ? 32 : 21);
        }
        var expenses = new[] {
            (Day: 2, Description: "Client workshop travel", Category: "Travel", Amount: 184.50),
            (Day: 4, Description: "Training materials", Category: "Training", Amount: 62.00),
            (Day: 7, Description: "Project meeting", Category: "Meals", Amount: 48.75),
            (Day: 9, Description: "Replacement headset", Category: "Equipment", Amount: 89.90)
        };
        for (int index = 0; index < expenses.Length; index++) {
            int row = index + 5;
            var expense = expenses[index];
            sheet.Cell(row, 1, new DateTime(2026, 9, expense.Day), numberFormat: "dd mmm yyyy");
            sheet.Cell(row, 2, expense.Description);
            sheet.Cell(row, 3, expense.Category);
            sheet.Cell(row, 4, expense.Amount, numberFormat: "#,##0.00");
            sheet.Cell(row, 5, index < 2 ? "Approved" : "Submitted");
        }
        sheet.AddTable("A4:E8", true, "Expenses", ExcelTableStyle.TableStyleMedium2);
        sheet.ValidationList("C5:C8", new[] { "Travel", "Training", "Meals", "Equipment" });
        sheet.ValidationList("E5:E8", new[] { "Submitted", "Approved", "Returned" });
        sheet.Cell(10, 3, "Total");
        sheet.CellFormula(10, 4, "SUM(D5:D8)");
        sheet.FormatCell(10, 4, "#,##0.00");
        sheet.CellBold(10, 4, true);
        sheet.Freeze(topRows: 4);
        document.Calculate();
        sheet.Range("A1:E12").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
