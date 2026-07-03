using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
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
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                int a1 = allText.IndexOf("A1", StringComparison.Ordinal);
                int b1 = allText.IndexOf("B1", StringComparison.Ordinal);
                int a2 = allText.IndexOf("A2", StringComparison.Ordinal);
                int b2 = allText.IndexOf("B2", StringComparison.Ordinal);
                Assert.True(a1 >= 0 && b1 > a1 && a2 > b1 && b2 > a2);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraphs_Headings_And_Tables() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeContent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeContent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native heading").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("First native paragraph");
                document.AddParagraph("Second native paragraph").SetBold().SetItalic();
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "N-A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "N-B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "N-A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "N-B2";
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native heading", allText);
                int first = allText.IndexOf("First native paragraph", StringComparison.Ordinal);
                int second = allText.IndexOf("Second native paragraph", StringComparison.Ordinal);
                int a1 = allText.IndexOf("N-A1", StringComparison.Ordinal);
                int b1 = allText.IndexOf("N-B1", StringComparison.Ordinal);
                int a2 = allText.IndexOf("N-A2", StringComparison.Ordinal);
                int b2 = allText.IndexOf("N-B2", StringComparison.Ordinal);
                Assert.True(first >= 0 && second > first);
                Assert.True(a1 >= 0 && b1 > a1 && a2 > b1 && b2 > a2);
            }
        }
    }
}
