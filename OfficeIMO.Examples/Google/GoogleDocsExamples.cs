using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Google {
    internal static class GoogleDocsExamples {
        public static void Example_Plan(string folderPath) {
            Console.WriteLine("[*] Google Docs - Building translation plan");

            string filePath = Path.Combine(folderPath, "GoogleDocs.Plan.docx");

            using var document = WordDocument.Create(filePath);
            document.AddParagraph("OfficeIMO Google Docs plan").SetStyle(WordParagraphStyles.Heading1);
            document.AddParagraph("This example builds a translation plan before any network calls happen.");
            document.Save();

            var plan = document.CreateGoogleDocsTranslationPlan(new GoogleDocsSaveOptions {
                Title = "OfficeIMO Google Docs plan"
            });

            foreach (var notice in plan.Report.Notices) {
                Console.WriteLine($"  {notice.Severity}: {notice.Feature} - {notice.Message}");
            }
        }

        public static async Task Example_ExportAsync(string folderPath) {
            Console.WriteLine("[*] Google Docs - Exporting document");

            var session = GoogleWorkspaceExampleSession.TryCreateSession();
            if (session == null) {
                GoogleWorkspaceExampleSession.PrintMissingTokenMessage();
                return;
            }

            string filePath = Path.Combine(folderPath, "GoogleDocs.Export.docx");

            using var document = WordDocument.Create(filePath);
            document.AddParagraph("OfficeIMO Google Docs export").SetStyle(WordParagraphStyles.Heading1);
            document.AddParagraph("This export uses session defaults for folder placement and retry behavior.");
            document.Save();

            try {
                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "OfficeIMO Google Docs export",
                    Location = new OfficeIMO.GoogleWorkspace.GoogleDriveFileLocation {
                        FolderId = session.Options.DefaultFolderId,
                        DriveId = session.Options.DefaultDriveId,
                    },
                });

                Console.WriteLine($"  DocumentId: {result.DocumentId}");
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
