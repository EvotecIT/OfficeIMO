using System;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates wrapping text with and without auto-fit behaviour so the differences are visible.
    /// </summary>
    public static class WrapText {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Wrap text scenarios");
            string filePath = System.IO.Path.Combine(folderPath, "WrapText.xlsx");

            using var document = ExcelDocument.Create(filePath);

            // Sheet 1: wrap enabled but no auto-fit applied
            var wrapOnly = document.AddWorkSheet("WrapOnly");
            wrapOnly.CellValue(1, 1, "Wrap Text (no auto-fit)");
            wrapOnly.CellValue(3, 1, "Line1\nLine2\nLine3");
            wrapOnly.WrapCells(3, 3, 1, 28);
            wrapOnly.CellValue(5, 1, "Very long sentence that will wrap once the row height is increased manually or via API.");
            wrapOnly.WrapCells(5, 5, 1, 28);

            // Sheet 2: wrap text plus AutoFitRows so the height adjusts
            var wrapAutoRows = document.AddWorkSheet("Wrap+AutoFitRows");
            wrapAutoRows.CellValue(1, 1, "Wrap Text + AutoFitRows");
            wrapAutoRows.CellValue(3, 1, "Line1\nLine2\nLine3");
            wrapAutoRows.WrapCells(3, 3, 1, 28);
            wrapAutoRows.CellValue(5, 1, "Very long sentence that will wrap onto multiple lines once row heights are adjusted.");
            wrapAutoRows.WrapCells(5, 5, 1, 28);
            wrapAutoRows.CellValue(7, 1, "Table note");
            wrapAutoRows.CellValue(8, 1, "Long description that benefits from wrapping to keep the table narrow and readable.");
            wrapAutoRows.CellValue(8, 2, "Short note");
            wrapAutoRows.WrapCells(8, 8, 1, 28);
            wrapAutoRows.AutoFitRows();

            // Sheet 3: wrap text combined with both AutoFitRows and AutoFitColumns to illustrate Excel's native behaviour
            var wrapAll = document.AddWorkSheet("Wrap+AutoFitAll");
            wrapAll.CellValue(1, 1, "Wrap Text + AutoFitRows + AutoFitColumns");
            wrapAll.CellValue(3, 1, "Line1\nLine2\nLine3");
            wrapAll.WrapCells(3, 3, 1, 28);
            wrapAll.CellValue(5, 1, "Very long sentence that demonstrates how Excel widens columns when auto-fit runs with wrap enabled.");
            wrapAll.WrapCells(5, 5, 1, 28);
            wrapAll.CellValue(7, 1, "Auto-fit note");
            wrapAll.CellValue(8, 1, "Auto-fitting columns widens them to the longest line even when wrap is enabled, matching Excel's UI.");
            wrapAll.CellValue(8, 2, "Short note");
            wrapAll.WrapCells(8, 8, 1, 28);
            wrapAll.AutoFitColumns();
            wrapAll.AutoFitRows();

            // Sheet 4: wrap column stays pinned while neighbours auto-fit
            var selective = document.AddWorkSheet("Wrap+SelectiveAutoFit");
            selective.CellValue(1, 1, "ID");
            selective.CellValue(1, 2, "Summary");
            selective.CellValue(1, 3, "Notes");
            for (int r = 2; r <= 8; r++) {
                selective.CellValue(r, 1, $"Row {r:00}");
                selective.CellValue(r, 2, "Pinned wrap column stays narrow even if we auto-fit other columns later on.");
                selective.CellValue(r, 3, "This column uses auto-fit so the text decides its width.");
            }
            selective.WrapCells(2, 8, 2, 26);
            selective.AutoFitColumnsExcept(new[] { 2 });

            // Sheet 5: SheetComposer example with column sizing helper
            var composer = new SheetComposer(document, "ComposerSizing");
            composer.Sheet.CellValue(1, 1, "Wrap");
            composer.Sheet.CellValue(1, 2, "AutoFit");
            composer.Sheet.CellValue(1, 3, "Number");
            for (int r = 2; r <= 6; r++) {
                composer.Sheet.CellValue(r, 1, "Long description that should stay narrow.");
                composer.Sheet.CellValue(r, 2, "Another sentence that can auto-fit to whatever width it needs.");
                // Explicit cast to avoid decimal/double overload ambiguity
                composer.Sheet.CellValue(r, 3, (double)(r * 123));
            }
            composer.ApplyColumnSizing("A1:C6", opts => {
                opts.WrapHeaders.Add("Wrap");
                opts.WrapWidth = 24;
                opts.AutoFitHeaders.Add("AutoFit");
                opts.NumericHeaders.Add("Number");
            });

            document.Save(openExcel);
        }
    }
}
