using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
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
                redRun.ColorHex = "FF0000";
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
        public void SaveAsPdf_OfficeIMOEngine_PortableDeterministicPolicyDoesNotEmbedSystemFonts() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontNoEmbedding.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontNoEmbedding.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Settings.FontFamily = "Calibri";
                document.AddParagraph("Document default font should not embed host fonts by default.");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
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
                    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
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
        public void SaveAsPdf_OfficeIMOEngine_Embeds_Distinct_Mapped_Run_Font_In_A_Separate_Slot() {
            if (!PdfEmbeddedFontFamily.TryFromSystem("Calibri", out PdfEmbeddedFontFamily? _) ||
                !PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfEmbeddedFontFamily? _)) {
                return;
            }

            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDistinctMappedRunFonts.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDistinctMappedRunFonts.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Settings.FontFamily = "Calibri";
                document.AddParagraph("Calibri default");
                document.AddParagraph().AddText("Arial run").SetFontFamily("Arial");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));

            Assert.Contains("/BaseFont /Calibri", content, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /Arial", content, StringComparison.Ordinal);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Reports_When_Distinct_Mapped_Font_Slots_Are_Exhausted() {
            IReadOnlyList<string> fontFamilies = FindMappedEmbeddableSansFontFamilies(4);
            if (fontFamilies.Count < 4) {
                return;
            }

            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMappedFontSlotExhaustion.docx");
            using WordDocument document = WordDocument.Create(docPath);
            document.Settings.FontFamily = fontFamilies[0];
            for (int index = 0; index < fontFamilies.Count; index++) {
                document.AddParagraph().AddText("Distinct font " + index).SetFontFamily(fontFamilies[index]);
            }

            document.Save();
            PdfDocumentConversionResult result = document.ToPdfDocumentResult(new PdfSaveOptions {
                IncludePageNumbers = false
            });

            PdfConversionWarning[] warnings = result.Warnings
                .Where(item => item.Code == "NativeFontFamilySlotExhausted")
                .ToArray();
            Assert.NotEmpty(warnings);
            Assert.Contains(warnings, warning => warning.Details["fontFamily"] == fontFamilies[fontFamilies.Count - 1]);
            Assert.All(warnings, warning => Assert.Contains(warning.Details["fontFamily"], fontFamilies));
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Embeds_Unmapped_Run_Font_When_Allowed_And_Available() {
            string? fontFamily = FindUnmappedEmbeddableWordFontFamily();
            if (fontFamily == null) {
                return;
            }

            const string unicodeMarker = "RichFontUnicode Zażółć";
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeUnmappedRunFontEmbedding.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeUnmappedRunFontEmbedding.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                paragraph.AddText(unicodeMarker).SetFontFamily(fontFamily);
                paragraph.AddText(" After");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains(unicodeMarker, pageText, StringComparison.Ordinal);
            }

            Assert.Contains("/Subtype /Type0", content, StringComparison.Ordinal);
            Assert.Contains("/Encoding /Identity-H", content, StringComparison.Ordinal);
            Assert.Contains("/ToUnicode", content, StringComparison.Ordinal);
            Assert.Contains("/FontFile2", content, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /" + SanitizeExpectedPdfFontName(fontFamily), content, StringComparison.Ordinal);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Embeds_Unmapped_Table_Cell_Run_Font_When_Allowed_And_Available() {
            string? fontFamily = FindUnmappedEmbeddableWordFontFamily();
            if (fontFamily == null) {
                return;
            }

            const string cellMarker = "TableCellRichFont";
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeUnmappedTableCellFontEmbedding.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeUnmappedTableCellFontEmbedding.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(1, 1);
                WordParagraph cellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
                cellParagraph.Text = string.Empty;
                cellParagraph.AddText(cellMarker).SetFontFamily(fontFamily);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains(cellMarker, pageText, StringComparison.Ordinal);
            }

            Assert.Contains("/Subtype /Type0", content, StringComparison.Ordinal);
            Assert.Contains("/FontFile2", content, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /" + SanitizeExpectedPdfFontName(fontFamily), content, StringComparison.Ordinal);
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
        public void SaveAsPdf_OfficeIMOEngine_Maps_Style_Baseline_Run_Properties() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStyleBaselineRun.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStyleBaselineRun.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Superscript Character Style" },
                    new StyleRunProperties(
                        new FontSize { Val = "40" },
                        new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }))
                {
                    Type = StyleValues.Character,
                    StyleId = "NativeSuperscriptCharacterStyle",
                    CustomStyle = true
                });
                styles.Append(new Style(
                    new StyleName { Val = "Native Subscript Paragraph Style" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(
                        new FontSize { Val = "40" },
                        new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "NativeSubscriptParagraphStyle",
                    CustomStyle = true
                });

                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                paragraph.AddText("StyledSuper").SetCharacterStyleId("NativeSuperscriptCharacterStyle");
                paragraph.AddText(" After");

                WordParagraph styledParagraph = document.AddParagraph("StyledSub");
                styledParagraph.SetStyleId("NativeSubscriptParagraphStyle");

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

                Assert.Equal(1, CountOccurrences(pageText, "StyledSuper"));
                Assert.Equal(1, CountOccurrences(pageText, "StyledSub"));
            }

            Assert.Matches(@"7\s+Ts", content);
            Assert.Matches(@"-3\.6\s+Ts", content);
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

        private static string? FindUnmappedEmbeddableWordFontFamily() {
            foreach (string candidate in new[] {
                "Comic Sans MS",
                "Candara",
                "Corbel",
                "Constantia",
                "Franklin Gothic Medium",
                "Trebuchet MS",
                "Segoe Print",
                "Bahnschrift",
                "Book Antiqua"
            }) {
                if (!PdfStandardFontMapper.TryMapFontFamily(candidate, out _) &&
                    PdfEmbeddedFontFamily.TryFromSystem(candidate, out PdfEmbeddedFontFamily? family) &&
                    family != null) {
                    return family.FamilyName;
                }
            }

            return null;
        }

        private static IReadOnlyList<string> FindMappedEmbeddableSansFontFamilies(int count) {
            var result = new List<string>(count);
            foreach (string candidate in new[] {
                "Arial",
                "Calibri",
                "Aptos",
                "Segoe UI",
                "Tahoma",
                "Verdana"
            }) {
                if (PdfStandardFontMapper.TryMapFontFamily(candidate, out PdfStandardFont mapped) &&
                    PdfStandardFontMapper.GetFontFamily(mapped) == PdfStandardFont.Helvetica &&
                    PdfEmbeddedFontFamily.TryFromSystem(candidate, out PdfEmbeddedFontFamily? family) &&
                    family != null) {
                    result.Add(candidate);
                    if (result.Count == count) {
                        break;
                    }
                }
            }

            return result;
        }

        private static string SanitizeExpectedPdfFontName(string fontFamily) {
            var builder = new StringBuilder(fontFamily.Length + "-Regular".Length);
            foreach (char ch in fontFamily + "-Regular") {
                if (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' || ch == '+') {
                    builder.Append(ch);
                }
            }

            return builder.Length == 0 ? "OfficeIMOEmbeddedFont" : builder.ToString();
        }
    }
}
