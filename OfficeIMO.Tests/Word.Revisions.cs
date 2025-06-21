using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AcceptRevisions_RemovesTrackedChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChanges.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Before");

                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains(document._document.Body.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.Contains(document._document.Body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");

                document.AcceptRevisions();

                Assert.DoesNotContain(document._document.Body.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.DoesNotContain(document._document.Body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");
                Assert.Contains(document.Paragraphs, p => p.Text == "Before");
                Assert.Contains(document.Paragraphs, p => p.Text == "Added");
            }
        }

        [Fact]
        public void Test_RejectRevisions_RemovesInsertions() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesReject.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.RejectRevisions();
                Assert.DoesNotContain(document._document.Body.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.DoesNotContain(document._document.Body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");
                Assert.Contains(document.Paragraphs, p => p.Text == "Removed");
            }
        }
    }
}
