using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates wrapping text with and without auto-fit behaviour so the differences are visible.
    /// </summary>
    public static class WrapText
    {
        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Wrap text scenarios");
            string filePath = System.IO.Path.Combine(folderPath, "WrapText.xlsx");
            using var document = ExcelDocument.Create(filePath);
            // Sheet 1: wrap enabled but no auto-fit applied
            var wrapOnly = document.AddSheet("WrapOnly");
            wrapOnly.Cell(1, 1).SetValue("Wrap Text (no auto-fit)");
            wrapOnly.Cell(3, 1).SetValue("Line1\nLine2\nLine3");
            wrapOnly.Range(3, 1, 3, 1).SetWrapText(); wrapOnly.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapOnly.Cell(5, 1).SetValue("Very long sentence that will wrap once the row height is increased manually or via API.");
            wrapOnly.Range(5, 1, 5, 1).SetWrapText(); wrapOnly.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            // Sheet 2: wrap text plus AutoFitRows so the height adjusts
            var wrapAutoRows = document.AddSheet("Wrap+AutoFitRows");
            wrapAutoRows.Cell(1, 1).SetValue("Wrap Text + AutoFitRows");
            wrapAutoRows.Cell(3, 1).SetValue("Line1\nLine2\nLine3");
            wrapAutoRows.Range(3, 1, 3, 1).SetWrapText(); wrapAutoRows.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapAutoRows.Cell(5, 1).SetValue("Very long sentence that will wrap onto multiple lines once row heights are adjusted.");
            wrapAutoRows.Range(5, 1, 5, 1).SetWrapText(); wrapAutoRows.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapAutoRows.Cell(7, 1).SetValue("Table note");
            wrapAutoRows.Cell(8, 1).SetValue("Long description that benefits from wrapping to keep the table narrow and readable.");
            wrapAutoRows.Cell(8, 2).SetValue("Short note");
            wrapAutoRows.Range(8, 1, 8, 1).SetWrapText(); wrapAutoRows.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapAutoRows.AutoFitRows();
            // Sheet 3: wrap text combined with both AutoFitRows and AutoFitColumns to illustrate Excel's native behaviour
            var wrapAll = document.AddSheet("Wrap+AutoFitAll");
            wrapAll.Cell(1, 1).SetValue("Wrap Text + AutoFitRows + AutoFitColumns");
            wrapAll.Cell(3, 1).SetValue("Line1\nLine2\nLine3");
            wrapAll.Range(3, 1, 3, 1).SetWrapText(); wrapAll.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapAll.Cell(5, 1).SetValue("Very long sentence that demonstrates how Excel widens columns when auto-fit runs with wrap enabled.");
            wrapAll.Range(5, 1, 5, 1).SetWrapText(); wrapAll.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapAll.Cell(7, 1).SetValue("Auto-fit note");
            wrapAll.Cell(8, 1).SetValue("Auto-fitting columns widens them to the longest line even when wrap is enabled, matching Excel's UI.");
            wrapAll.Cell(8, 2).SetValue("Short note");
            wrapAll.Range(8, 1, 8, 1).SetWrapText(); wrapAll.SetColumnWidth(1, ExcelColumnWidth.Characters(28));
            wrapAll.AutoFitColumns();
            wrapAll.AutoFitRows();
            // Sheet 4: wrap column stays pinned while neighbours auto-fit
            var selective = document.AddSheet("Wrap+SelectiveAutoFit");
            selective.Cell(1, 1).SetValue("ID");
            selective.Cell(1, 2).SetValue("Summary");
            selective.Cell(1, 3).SetValue("Notes");
            for (int r = 2; r <= 8; r++)
            {
                selective.Cell(r, 1).SetValue($"Row {r:00}");
                selective.Cell(r, 2).SetValue("Pinned wrap column stays narrow even if we auto-fit other columns later on.");
                selective.Cell(r, 3).SetValue("This column uses auto-fit so the text decides its width.");
            }

            selective.Range(2, 2, 8, 2).SetWrapText(); selective.SetColumnWidth(2, ExcelColumnWidth.Characters(26));
            selective.AutoFitColumnsExcept(new[] { 2 });
            // Sheet 5: ExcelReportBuilder example with column sizing helper
            var composer = document.AddSheet("ComposerSizing", ExcelSheetNameValidationMode.Sanitize).Report();
            composer.Sheet.Cell(1, 1).SetValue("Wrap");
            composer.Sheet.Cell(1, 2).SetValue("AutoFit");
            composer.Sheet.Cell(1, 3).SetValue("Number");
            for (int r = 2; r <= 6; r++)
            {
                composer.Sheet.Cell(r, 1).SetValue("Long description that should stay narrow.");
                composer.Sheet.Cell(r, 2).SetValue("Another sentence that can auto-fit to whatever width it needs.");
                // Explicit cast to avoid decimal/double overload ambiguity
                composer.Sheet.Cell(r, 3).SetValue((double)(r * 123));
            }

            composer.ApplyColumnSizing("A1:C6", opts =>
            {
                opts.WrapHeaders.Add("Wrap");
                opts.WrapWidth = ExcelColumnWidth.Characters(24);
                opts.AutoFitHeaders.Add("AutoFit");
                opts.NumericHeaders.Add("Number");
            });
            document.Save();
            if (openExcel)
                document.OpenInApplication();
        }
    }
}