using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareDetectsInsertedText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_insert.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_insert.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Hello World");
                doc.Save(false);
            }

            string resultPath;
            using (WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
                resultPath = result.FilePath;
            }

            using WordprocessingDocument word = WordprocessingDocument.Open(resultPath, false);
            InsertedRun ins = word.MainDocumentPart.Document.Body.Descendants<InsertedRun>().FirstOrDefault();
            Assert.NotNull(ins);
            Assert.Equal(" World", ins.InnerText);
        }

        [Fact]
        public void CompareDetectsDeletedText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_delete.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello World");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_delete.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            string resultPath;
            using (WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
                resultPath = result.FilePath;
            }

            using WordprocessingDocument word = WordprocessingDocument.Open(resultPath, false);
            DeletedRun del = word.MainDocumentPart.Document.Body.Descendants<DeletedRun>().FirstOrDefault();
            Assert.NotNull(del);
            Assert.Equal(" World", del.InnerText);
        }

        [Fact]
        public void CompareDetectsFormattingChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_format.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_format.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                var p = doc.AddParagraph("Hello");
                p.Bold = true;
                doc.Save(false);
            }

            string resultPath;
            using (WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
                resultPath = result.FilePath;
            }

            using WordprocessingDocument word = WordprocessingDocument.Open(resultPath, false);
            Run run = word.MainDocumentPart.Document.Body.Descendants<Run>().First();
            Assert.NotNull(run.RunProperties);
            Assert.NotNull(run.RunProperties.RunPropertiesChange);
        }

        [Fact]
        public void CompareDetectsInsertedTableRow() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_row_insert.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_row_insert.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(2, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("Row2");
                doc.Save(false);
            }

            string resultPath;
            using (WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
                resultPath = result.FilePath;
            }

            using WordprocessingDocument word = WordprocessingDocument.Open(resultPath, false);
            InsertedRun ins = word.MainDocumentPart.Document.Body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "Row2");
            Assert.NotNull(ins);
        }

        [Fact]
        public void CompareDetectsDeletedTableRow() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_row_delete.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(2, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("Row2");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_row_delete.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                doc.Save(false);
            }

            string resultPath;
            using (WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
                resultPath = result.FilePath;
            }

            using WordprocessingDocument word = WordprocessingDocument.Open(resultPath, false);
            DeletedRun del = word.MainDocumentPart.Document.Body.Descendants<DeletedRun>().FirstOrDefault(r => r.InnerText == "Row2");
            Assert.NotNull(del);
        }

        [Fact]
        public void CompareDetectsImageReplacement() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "snail.bmp"));
                doc.Save(false);
            }

            string resultPath;
            using (WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
                resultPath = result.FilePath;
            }

            using WordprocessingDocument word = WordprocessingDocument.Open(resultPath, false);
            InsertedRun ins = word.MainDocumentPart.Document.Body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "[Image]");
            DeletedRun del = word.MainDocumentPart.Document.Body.Descendants<DeletedRun>().FirstOrDefault(r => r.InnerText == "[Image]");
            Assert.NotNull(ins);
            Assert.NotNull(del);
        }

    }
}
