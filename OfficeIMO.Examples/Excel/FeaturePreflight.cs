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
                ExcelSheet sheet = document.AddSheet("Report");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Score");
                sheet.Cell(2, 1).SetValue("Alpha");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("Beta");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.Cell(4, 2).SetFormula("SUM(B2:B3)");
                document.Calculate();
                document.Save();
                if (openExcel) document.OpenInApplication();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
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
