using System;
using System.IO;
using System.Text;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static class ReviewReports {
        internal static void Example_ReviewReportWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word review report workflow");

            string documentPath = Path.Combine(folderPath, "ReviewReportWorkflow.docx");
            using (WordDocument document = WordDocument.Create(documentPath)) {
                document.AddParagraph("Service Agreement Review").Style = WordParagraphStyles.Heading1;

                WordParagraph commentTarget = document.AddParagraph("The supplier must notify the customer within 48 hours.");
                commentTarget.AddComment("Legal Reviewer", "LR", "Please align this notification period with the incident policy.");
                WordComment comment = WordComment.GetAllComments(document)[^1];
                comment.AddReply("Document Owner", "DO", "Updated in the tracked revision below.");
                comment.MarkResolved();

                WordParagraph tracked = document.AddParagraph("Tracked clause: ");
                tracked.AddDeletedText("best effort", "Legal Reviewer", new DateTime(2026, 6, 1, 9, 0, 0, DateTimeKind.Utc));
                tracked.AddInsertedText("commercially reasonable efforts", "Legal Reviewer", new DateTime(2026, 6, 1, 9, 5, 0, DateTimeKind.Utc));

                document.Save(new WordSaveOptions { OpenAfterSave = openWord });
            }

            using (WordDocument document = WordDocument.Load(documentPath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                WordReviewReport report = document.InspectReviewReport();
                File.WriteAllText(Path.Combine(folderPath, "ReviewReportWorkflow.md"), report.ToMarkdown(), Encoding.UTF8);
                File.WriteAllText(Path.Combine(folderPath, "ReviewReportWorkflow.json"), report.ToJson(), Encoding.UTF8);
            }
        }
    }
}
