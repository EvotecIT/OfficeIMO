using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates configuring worksheet validation diagnostics for parallel writes.
    /// </summary>
    public static class WorksheetValidation {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Configure worksheet validation diagnostics");
            string filePath = System.IO.Path.Combine(folderPath, "WorksheetValidationDiagnostics.xlsx");

            using var document = ExcelDocument.Create(filePath);

            document.Execution.WorksheetValidation = WorksheetValidationMode.DiagnosticsOnly;
            document.Execution.DiagnosticsRequested = true;
            document.Execution.OnInfo = message => Console.WriteLine($"[diagnostic] {message}");

            var sheet = document.AddWorkSheet("Telemetry");
            var cells = Enumerable.Range(1, 20)
                .SelectMany(row => Enumerable.Range(1, 3).Select(col => (row, col, (object)$"R{row}C{col}")))
                .ToList();

            sheet.CellValuesParallel(cells);

            // Switch to disabled mode for throughput-sensitive sections.
            document.Execution.WorksheetValidation = WorksheetValidationMode.Disabled;
            var fastCells = Enumerable.Range(21, 20).Select(i => (i, 1, (object)$"Fast {i}"));
            sheet.CellValuesParallel(fastCells);

            document.Save(openExcel);
        }
    }
}
