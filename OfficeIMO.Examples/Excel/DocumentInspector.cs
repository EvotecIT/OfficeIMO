using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>Removes comments, hidden sheets and personal metadata from a workbook before sharing.</summary>
    public static class DocumentInspector {
        public static void Example_PrepareForSharing(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Remove personal information and hidden sheets before sharing");
            string filePath = Path.Combine(folderPath, "Document Inspector.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet summary = document.AddSheet("Summary");
                summary.Cell(1, 1).SetValue("Revenue");
                summary.SetComment(1, 1, "Numbers are preliminary.", "Jane Doe");
                ExcelSheet workings = document.AddSheet("Workings");
                workings.Cell(1, 1).SetValue("Internal margin");
                workings.SetHidden(true);
                document.Properties.Creator = "Jane Doe";
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                OfficeHiddenDataReport report = document.RemoveHiddenData(new OfficeHiddenDataOptions {
                    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.HiddenContent
                });
                foreach (OfficeHiddenDataFinding finding in report.Findings) {
                    Console.WriteLine($"    {finding.Outcome} {finding.Kind} at {finding.Location}: {finding.Description}");
                }
                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}
