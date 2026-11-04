using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Google {
    internal static class GoogleSheetsExamples {
        public static void Example_Plan(string folderPath) {
            Console.WriteLine("[*] Google Sheets - Building translation plan");

            string filePath = Path.Combine(folderPath, "GoogleSheets.Plan.xlsx");

            using var workbook = ExcelDocument.Create(filePath);
            var summary = workbook.AddSheet("Summary");
            summary.Cell(1, 1).SetValue("Quarter");
            summary.Cell(1, 2).SetValue("Revenue");
            summary.Cell(2, 1).SetValue("Q1");
            summary.Cell(2, 2).SetValue(125000);
            summary.Cell(3, 1).SetValue("Q2");
            summary.Cell(3, 2).SetValue(151500);
            workbook.Save();

            var plan = workbook.BuildGoogleSheetsPlan(new GoogleSheetsSaveOptions {
                Title = "OfficeIMO Google Sheets plan"
            });

            foreach (var notice in plan.Report.Notices) {
                Console.WriteLine($"  {notice.Severity}: {notice.Feature} - {notice.Message}");
            }
        }

        public static async Task Example_ExportAsync(string folderPath) {
            Console.WriteLine("[*] Google Sheets - Exporting workbook");

            GoogleWorkspaceSession? session = GoogleWorkspaceExampleSession.TryCreateSession();
            if (session == null) {
                GoogleWorkspaceExampleSession.PrintMissingTokenMessage();
                return;
            }

            string filePath = Path.Combine(folderPath, "GoogleSheets.Export.xlsx");

            using var workbook = ExcelDocument.Create(filePath);
            var summary = workbook.AddSheet("Summary");
            summary.Cell(1, 1).SetValue("Quarter");
            summary.Cell(1, 2).SetValue("Revenue");
            summary.Cell(2, 1).SetValue("Q3");
            summary.Cell(2, 2).SetValue(178000);
            summary.Cell(3, 1).SetValue("Q4");
            summary.Cell(3, 2).SetValue(210250);
            workbook.Save();

            try {
                var result = await workbook.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "OfficeIMO Google Sheets export",
                    Location = new GoogleDriveFileLocation {
                        FolderId = session.Options.DefaultFolderId,
                        DriveId = session.Options.DefaultDriveId,
                    },
                });

                Console.WriteLine($"  SpreadsheetId: {result.SpreadsheetId}");
                Console.WriteLine($"  WebViewLink: {result.WebViewLink}");

                foreach (var notice in result.Report.Notices) {
                    Console.WriteLine($"  {notice.Severity}: {notice.Feature} - {notice.Message}");
                }
            } catch (GoogleWorkspaceExportException exception) {
                GoogleWorkspaceExampleSession.PrintExportFailure(exception);
            }
        }
    }
}
