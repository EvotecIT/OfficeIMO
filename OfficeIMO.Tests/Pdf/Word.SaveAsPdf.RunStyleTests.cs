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
        public void SaveAsPdf_OfficeIMOEngine_Does_Not_Leak_Run_Color_To_Following_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRunColorReset.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRunColorReset.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                WordParagraph redRun = paragraph.AddText("Red");
                redRun.ColorHex = "ff0000";
                paragraph.AddText("After");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int redText = content.IndexOf("<526564>", StringComparison.Ordinal);
            int afterText = content.IndexOf("<4166746572>", StringComparison.Ordinal);

            Assert.True(redText >= 0, "Expected encoded 'Red' text in the generated PDF content stream.");
            Assert.True(afterText > redText, "Expected encoded 'After' text after the red Word run.");

            int redColorBeforeRed = content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal);
            int blackColorBeforeAfter = content.LastIndexOf("0 0 0 rg", afterText, StringComparison.Ordinal);
            int redColorBeforeAfter = content.LastIndexOf("1 0 0 rg", afterText, StringComparison.Ordinal);

            Assert.True(redColorBeforeRed >= 0, "Expected the Word run color to emit a red PDF fill color.");
            Assert.True(blackColorBeforeAfter > redText, "Expected the following uncolored Word run to reset to black/default PDF fill color.");
            Assert.True(redColorBeforeAfter < blackColorBeforeAfter, "Expected the following Word run not to inherit the previous red fill color.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Run_And_Paragraph_Font_Sizes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRunFontSizes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRunFontSizes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("ParagraphSized").SetFontSize(16);

                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Small");
                paragraph.AddText("Large").SetFontSize(18);
                paragraph.AddText("Normal");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int paragraphText = content.IndexOf("<50617261677261706853697A6564>", StringComparison.Ordinal);
            int smallText = content.IndexOf("<536D616C6C>", StringComparison.Ordinal);
            int largeText = content.IndexOf("<4C61726765>", StringComparison.Ordinal);
            int normalText = content.IndexOf("<4E6F726D616C>", StringComparison.Ordinal);

            Assert.True(paragraphText >= 0, "Expected encoded 'ParagraphSized' text in the native PDF content stream.");
            Assert.True(smallText > paragraphText, "Expected encoded 'Small' text after the paragraph-sized paragraph.");
            Assert.True(largeText > smallText, "Expected encoded 'Large' text after the default-sized run.");
            Assert.True(normalText > largeText, "Expected encoded 'Normal' text after the large run.");

            int paragraphSizeBeforeParagraph = content.LastIndexOf("/F1 16 Tf", paragraphText, StringComparison.Ordinal);
            int defaultSizeBeforeSmall = content.LastIndexOf("/F1 11 Tf", smallText, StringComparison.Ordinal);
            int largeSizeBeforeLarge = content.LastIndexOf("/F1 18 Tf", largeText, StringComparison.Ordinal);
            int defaultSizeBeforeNormal = content.LastIndexOf("/F1 11 Tf", normalText, StringComparison.Ordinal);

            Assert.True(paragraphSizeBeforeParagraph >= 0, "Expected paragraph FontSize to emit a 16-point native PDF run.");
            Assert.True(defaultSizeBeforeSmall > paragraphText, "Expected the next paragraph to return to the default font size.");
            Assert.True(largeSizeBeforeLarge > smallText, "Expected the Word run font size to emit an 18-point native PDF run.");
            Assert.True(defaultSizeBeforeNormal > largeText, "Expected following Word runs not to inherit the previous run font size.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Embeds_Document_Default_Font_When_Available() {
            if (!PdfEmbeddedFontFamily.TryFromSystem("Calibri", out PdfEmbeddedFontFamily? _)) {
                return;
            }

            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontEmbedding.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontEmbedding.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Settings.FontFamily = "Calibri";
                document.AddParagraph("Document default font should be embedded for stable PDF viewers.");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);

            Assert.Contains("/Subtype /Type0", content, StringComparison.Ordinal);
            Assert.Contains("/Subtype /CIDFontType2", content, StringComparison.Ordinal);
            Assert.Contains("/Encoding /Identity-H", content, StringComparison.Ordinal);
            Assert.Contains("/FontFile2", content, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /Calibri", content, StringComparison.Ordinal);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Run_Highlight_To_Background() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRunHighlight.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRunHighlight.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                paragraph.AddText("Marked").SetHighlight(HighlightColorValues.Yellow);
                paragraph.AddText("After");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int markedText = content.IndexOf("<4D61726B6564>", StringComparison.Ordinal);
            int afterText = content.IndexOf("<4166746572>", StringComparison.Ordinal);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "Before"));
                Assert.Equal(1, CountOccurrences(pageText, "Marked"));
                Assert.Equal(1, CountOccurrences(pageText, "After"));
            }

            Assert.True(markedText >= 0, "Expected encoded 'Marked' text in the native PDF content stream.");
            Assert.True(afterText > markedText, "Expected encoded 'After' text after the highlighted Word run.");

            int highlightFill = content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal);
            int highlightRect = content.LastIndexOf(" re f", markedText, StringComparison.Ordinal);

            Assert.True(highlightFill >= 0, "Expected Word run highlight to emit a yellow PDF fill color.");
            Assert.True(highlightRect > highlightFill, "Expected Word run highlight to emit a filled rectangle behind the text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Document_Background_Color() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentBackground.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentBackground.pdf");
            string marker = "WordBackgroundPdfMarker";

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Background.SetColorHex("EAF4FF");
                document.AddParagraph(marker);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            string markerHex = BitConverter.ToString(Encoding.ASCII.GetBytes(marker)).Replace("-", string.Empty);
            int backgroundFill = content.IndexOf("0.918 0.957 1 rg", StringComparison.Ordinal);
            int backgroundRect = backgroundFill < 0 ? -1 : content.IndexOf(" re f", backgroundFill, StringComparison.Ordinal);
            int markerText = content.IndexOf("<" + markerHex + ">", StringComparison.Ordinal);

            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains(marker, pageText, StringComparison.Ordinal);
            }

            Assert.True(backgroundFill >= 0, "Expected Word document background color to emit a PDF page fill color.");
            Assert.True(backgroundRect > backgroundFill, "Expected Word document background color to emit a filled page rectangle.");
            Assert.True(markerText > backgroundRect, "Expected document background to render before Word paragraph text.");
        }
    }
}
