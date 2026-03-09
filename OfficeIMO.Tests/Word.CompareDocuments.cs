using System.IO;
using System.Linq;
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

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault();
            Assert.NotNull(ins);
            Assert.Equal(" World", ins!.InnerText);
            AssertNoTempArtifact(result);
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

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun? del = body.Descendants<DeletedRun>().FirstOrDefault();
            Assert.NotNull(del);
            Assert.Equal(" World", del!.InnerText);
            AssertNoTempArtifact(result);
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

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Run run = body.Descendants<Run>().First();
            Assert.NotNull(run.RunProperties);
            Assert.NotNull(run.RunProperties!.RunPropertiesChange);
            AssertNoTempArtifact(result);
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

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "Row2");
            Assert.NotNull(ins);
            AssertNoTempArtifact(result);
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

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun? del = body.Descendants<DeletedRun>().FirstOrDefault(r => r.InnerText == "Row2");
            Assert.NotNull(del);
            AssertNoTempArtifact(result);
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

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "[Image]");
            DeletedRun? del = body.Descendants<DeletedRun>().FirstOrDefault(r => r.InnerText == "[Image]");
            Assert.NotNull(ins);
            Assert.NotNull(del);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsInsertedTableCellInExistingRow() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_cell_insert.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_cell_insert.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Right");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Assert.Equal(2, result.Tables[0].Rows[0].CellsCount);

            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "Right");
            Assert.NotNull(ins);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void ComparePreservesListFormattingOnChangedParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_list_format.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                var list = doc.AddList(WordListStyle.Numbered);
                list.AddItem("Item 1");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_list_format.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                var list = doc.AddList(WordListStyle.Numbered);
                list.AddItem("Item 1 updated");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            WordParagraph? listParagraph = result.Paragraphs.FirstOrDefault(p => p.Text.Contains("Item 1"));

            Assert.NotNull(listParagraph);
            Assert.True(listParagraph!.IsListItem);
            AssertNoTempArtifact(result);
        }

        private static void AssertNoTempArtifact(WordDocument document) {
            Assert.Equal(string.Empty, document.FilePath);
        }
    }
}
