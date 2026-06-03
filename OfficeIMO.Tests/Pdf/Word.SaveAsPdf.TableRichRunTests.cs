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
            string content = Encoding.ASCII.GetString(bytes);
            int redText = content.IndexOf("<43656C6C526564>", StringComparison.Ordinal);
            int boldText = content.IndexOf("<43656C6C426F6C64>", StringComparison.Ordinal);
            int markedText = content.IndexOf("<43656C6C4D61726B6564>", StringComparison.Ordinal);
            int largeText = content.IndexOf("<43656C6C4C61726765>", StringComparison.Ordinal);

            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "CellPlain"));
                Assert.Equal(1, CountOccurrences(pageText, "CellRed"));
                Assert.Equal(1, CountOccurrences(pageText, "CellBold"));
                Assert.Equal(1, CountOccurrences(pageText, "CellMarked"));
                Assert.Equal(1, CountOccurrences(pageText, "CellLarge"));
            }

            Assert.True(redText >= 0, "Expected encoded 'CellRed' text in the native table PDF content stream.");
            Assert.True(boldText > redText, "Expected encoded 'CellBold' text after the colored table cell run.");
            Assert.True(markedText > boldText, "Expected encoded 'CellMarked' text after the bold table cell run.");
            Assert.True(largeText > markedText, "Expected encoded 'CellLarge' text after the highlighted table cell run.");
            Assert.True(content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal) >= 0, "Expected Word table cell run color to emit a red PDF fill color.");
            Assert.True(content.LastIndexOf("/F2 ", boldText, StringComparison.Ordinal) >= 0, "Expected Word table cell bold run to use the bold PDF font resource.");
            Assert.True(content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal) >= 0, "Expected Word table cell run highlight to emit a yellow PDF fill color.");
            Assert.True(content.LastIndexOf(" 18 Tf", largeText, StringComparison.Ordinal) >= 0, "Expected Word table cell run font size to emit an 18-point PDF run.");
        }
    }
}
