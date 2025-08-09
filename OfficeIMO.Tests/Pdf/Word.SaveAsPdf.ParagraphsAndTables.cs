using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_Renders_Paragraphs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfParagraphs.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfParagraphs.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("First paragraph");
                document.AddParagraph("Second paragraph");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                int first = allText.IndexOf("First paragraph", StringComparison.Ordinal);
                int second = allText.IndexOf("Second paragraph", StringComparison.Ordinal);
                Assert.True(first >= 0 && second > first);
            }
        }

        [Fact]
        public void SaveAsPdf_Renders_Tables() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfTableContent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfTableContent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                int a1 = allText.IndexOf("A1", StringComparison.Ordinal);
                int b1 = allText.IndexOf("B1", StringComparison.Ordinal);
                int a2 = allText.IndexOf("A2", StringComparison.Ordinal);
                int b2 = allText.IndexOf("B2", StringComparison.Ordinal);
                Assert.True(a1 >= 0 && b1 > a1 && a2 > b1 && b2 > a2);
            }
        }
    }
}