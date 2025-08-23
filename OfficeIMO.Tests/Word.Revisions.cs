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
                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);
                Assert.Contains(body!.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");

                document.AcceptRevisions();

                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");
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
                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);
                Assert.DoesNotContain(body!.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");
                Assert.Contains(document.Paragraphs, p => p.Text == "Removed");
            }
        }

        [Fact]
        public void Test_TrackedChanges_Validation() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesValidation.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");
                document.Save(false);

                var errors = document.ValidateDocument();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ConvertRevisionsToMarkup_PreservesTextWithFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesMarkup.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.ConvertRevisionsToMarkup();

                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);
                Assert.DoesNotContain(body!.Descendants<InsertedRun>(), r => r.InnerText == "Added");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), r => r.InnerText == "Removed");

                var insertedRun = body.Descendants<Run>().FirstOrDefault(r => r.InnerText == "Added");
                Assert.NotNull(insertedRun);
                Assert.NotNull(insertedRun!.RunProperties);
                Assert.NotNull(insertedRun.RunProperties!.Underline);
                Assert.Equal("0000FF", insertedRun.RunProperties.Color?.Val);

                var deletedRun = body.Descendants<Run>().FirstOrDefault(r => r.InnerText == "Removed");
                Assert.NotNull(deletedRun);
                Assert.NotNull(deletedRun!.RunProperties);
                Assert.NotNull(deletedRun.RunProperties!.Strike);
                Assert.Equal("FF0000", deletedRun.RunProperties.Color?.Val);
            }
        }
    }
}
