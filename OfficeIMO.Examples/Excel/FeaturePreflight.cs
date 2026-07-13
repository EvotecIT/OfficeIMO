using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates workflow-level feature preflight before read, edit, template, or PDF operations.
    /// </summary>
    public static class FeaturePreflight {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Feature preflight");
            string filePath = Path.Combine(folderPath, "FeaturePreflight.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "Beta");
                sheet.CellValue(3, 2, 20d);
                sheet.CellFormula(4, 2, "SUM(B2:B3)");
                document.Calculate();
                document.Save();
                if (openExcel) document.OpenInApplication();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                foreach (ExcelPreflightCapability capability in Enum.GetValues(typeof(ExcelPreflightCapability))) {
                    Console.WriteLine($"{capability}: {(report.Can(capability) ? "ready" : "blocked")}");
                    foreach (string diagnostic in report.GetCapabilityDiagnostics(capability)) {
                        Console.WriteLine("  - " + diagnostic);
                    }
                }
            }
        }
    }
}
