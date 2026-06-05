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
        public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_Rich_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRichRuns.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRichRuns.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(1, 1);
                WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
                paragraph.Text = string.Empty;
                paragraph.AddText("CellPlain ");
                WordParagraph red = paragraph.AddText("CellRed");
                red.ColorHex = "ff0000";
                paragraph.AddText(" ");
                paragraph.AddText("CellBold").SetBold();
                paragraph.AddText(" ");
                paragraph.AddText("CellMarked").SetHighlight(HighlightColorValues.Yellow);
                paragraph.AddText(" ");
                paragraph.AddText("CellLarge").SetFontSize(18);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);

            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "CellPlain"));
                Assert.Equal(1, CountOccurrences(pageText, "CellRed"));
                Assert.Equal(1, CountOccurrences(pageText, "CellBold"));
                Assert.Equal(1, CountOccurrences(pageText, "CellMarked"));
                Assert.Equal(1, CountOccurrences(pageText, "CellLarge"));
            }

            Assert.Contains("1 0 0 rg", content, StringComparison.Ordinal);
            Assert.Matches(@"/F\d+\s+(10|11)\s+Tf", content);
            Assert.Contains("1 1 0 rg", content, StringComparison.Ordinal);
            Assert.Matches(@"/F\d+\s+18\s+Tf", content);
        }
    }
}
