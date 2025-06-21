using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
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
                Assert.Equal(2, document.Paragraphs.Count);
                Assert.Equal("Before", document.Paragraphs[0].Text);
                Assert.Equal("Added", document.Paragraphs[1].Text);
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
                Assert.Single(document.Paragraphs);
                Assert.Equal("Removed", document.Paragraphs[0].Text);
            }
        }
    }
}
