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
            string content = ReadPdfPageContent(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "Before"));
                Assert.Equal(1, CountOccurrences(pageText, "Red"));
                Assert.Equal(1, CountOccurrences(pageText, "After"));
            }

            int redColor = content.IndexOf("1 0 0 rg", StringComparison.Ordinal);
            int blackColorAfterRed = content.IndexOf("0 0 0 rg", redColor + "1 0 0 rg".Length, StringComparison.Ordinal);

            Assert.True(redColor >= 0, "Expected the Word run color to emit a red PDF fill color.");
            Assert.True(blackColorAfterRed > redColor, "Expected the following uncolored Word run to reset to black/default PDF fill color.");
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
            string content = ReadPdfPageContent(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "ParagraphSized"));
                Assert.Equal(1, CountOccurrences(pageText, "Small"));
                Assert.Equal(1, CountOccurrences(pageText, "Large"));
                Assert.Equal(1, CountOccurrences(pageText, "Normal"));
            }

            Assert.Matches(@"/F\d+\s+16\s+Tf", content);
            Assert.Matches(@"/F\d+\s+18\s+Tf", content);
            Assert.True(Regex.Matches(content, @"/F\d+\s+11\s+Tf").Count >= 2, "Expected default-sized native PDF runs before and after explicit font sizes.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Does_Not_Embed_System_Fonts_By_Default() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontNoEmbedding.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontNoEmbedding.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Settings.FontFamily = "Calibri";
                document.AddParagraph("Document default font should not embed host fonts by default.");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);

            Assert.DoesNotContain("/FontFile2", content, StringComparison.Ordinal);
            Assert.DoesNotContain("/BaseFont /Calibri", content, StringComparison.Ordinal);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Embeds_Document_Default_Font_When_Allowed_And_Available() {
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
                    AllowSystemFontEmbedding = true,
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
            string content = ReadPdfPageContent(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "Before"));
                Assert.Equal(1, CountOccurrences(pageText, "Marked"));
                Assert.Equal(1, CountOccurrences(pageText, "After"));
            }

            int highlightFill = content.IndexOf("1 1 0 rg", StringComparison.Ordinal);
            int highlightRect = highlightFill < 0 ? -1 : content.IndexOf(" re f", highlightFill, StringComparison.Ordinal);

            Assert.True(highlightFill >= 0, "Expected Word run highlight to emit a yellow PDF fill color.");
            Assert.True(highlightRect > highlightFill, "Expected Word run highlight to emit a filled rectangle behind the text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Inherited_Character_Style_Run_Properties() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeInheritedCharacterStyleRun.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeInheritedCharacterStyleRun.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Base Character Style" },
                    new StyleRunProperties(
                        new Color { Val = "C00000" },
                        new FontSize { Val = "32" }))
                {
                    Type = StyleValues.Character,
                    StyleId = "NativeBaseCharacterStyle",
                    CustomStyle = true
                });
                styles.Append(new Style(
                    new StyleName { Val = "Native Derived Character Style" },
                    new BasedOn { Val = "NativeBaseCharacterStyle" },
                    new StyleRunProperties(new Underline { Val = UnderlineValues.Single }))
                {
                    Type = StyleValues.Character,
                    StyleId = "NativeDerivedCharacterStyle",
                    CustomStyle = true
                });

                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                paragraph.AddText("StyledChar").SetCharacterStyleId("NativeDerivedCharacterStyle");
                paragraph.AddText(" After");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "Before"));
                Assert.Equal(1, CountOccurrences(pageText, "StyledChar"));
                Assert.Equal(1, CountOccurrences(pageText, "After"));
            }

            Assert.Matches(@"/F\d+\s+16\s+Tf", content);
            Assert.Contains("0.753 0 0 rg", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Character_Style_Run_Properties_In_Table_Cells() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellCharacterStyleRun.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellCharacterStyleRun.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeTableCellCharacterStyle";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Table Cell Character Style" },
                    new StyleRunProperties(
                        new Color { Val = "C00000" },
                        new FontSize { Val = "36" }))
                {
                    Type = StyleValues.Character,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordTable table = document.AddTable(1, 1);
                WordParagraph cellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
                cellParagraph.Text = string.Empty;
                cellParagraph.AddText("CellStyledChar").SetCharacterStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 220),
                    Margins = PageMargins.Uniform(40),
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "CellStyledChar"));
            }

            Assert.Matches(@"/F\d+\s+18\s+Tf", content);
            Assert.Contains("0.753 0 0 rg", content);
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
            string content = ReadPdfPageContent(bytes);
            int backgroundFill = content.IndexOf("0.918 0.957 1 rg", StringComparison.Ordinal);
            int backgroundRect = backgroundFill < 0 ? -1 : content.IndexOf(" re f", backgroundFill, StringComparison.Ordinal);

            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains(marker, pageText, StringComparison.Ordinal);
            }

            Assert.True(backgroundFill >= 0, "Expected Word document background color to emit a PDF page fill color.");
            Assert.True(backgroundRect > backgroundFill, "Expected Word document background color to emit a filled page rectangle.");
        }
    }
}
