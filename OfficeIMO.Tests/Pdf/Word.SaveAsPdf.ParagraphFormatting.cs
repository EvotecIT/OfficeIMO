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
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Baseline_Formatting() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBaseline.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBaseline.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native baseline formatting");
                WordParagraph superscript = document.AddParagraph("Native paragraph superscript");
                superscript.FontSize = 20;
                superscript.SetSuperScript();
                WordParagraph subscript = document.AddParagraph("Native paragraph subscript");
                subscript.FontSize = 20;
                subscript.SetSubScript();
                WordParagraph mixed = document.AddParagraph();
                mixed.AddText("Native mixed baseline ");
                WordParagraph runSuperscript = mixed.AddText("run superscript");
                runSuperscript.FontSize = 20;
                runSuperscript.SetSuperScript();
                mixed.AddText(" ");
                WordParagraph runSubscript = mixed.AddText("run subscript");
                runSubscript.FontSize = 20;
                runSubscript.SetSubScript();
                document.AddParagraph("After native baseline formatting");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Before native baseline formatting", allText);
                Assert.Contains("Native paragraph superscript", allText);
                Assert.Contains("Native paragraph subscript", allText);
                Assert.Contains("Native mixed baseline", allText);
                Assert.Contains("run superscript", allText);
                Assert.Contains("run subscript", allText);
                Assert.Contains("After native baseline formatting", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("7 Ts", content);
            Assert.Contains("-3.6 Ts", content);
            Assert.Contains("0 Ts", content);
            Assert.Contains("/F1 13 Tf", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Text_Wrapping_Breaks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTextWrappingBreaks.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTextWrappingBreaks.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native text wrapping breaks");
                WordParagraph paragraph = document.AddParagraph("NativeSoftFirst");
                paragraph.AddBreak();
                paragraph.AddText("NativeSoftSecond");
                WordTable table = document.AddTable(1, 1);
                WordParagraph cellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
                cellParagraph.Text = string.Empty;
                cellParagraph.AddText("CellSoftFirst");
                cellParagraph.AddBreak();
                cellParagraph.AddText("CellSoftSecond");
                document.AddParagraph("After native text wrapping breaks");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("NativeSoftFirst", allText);
                Assert.Contains("NativeSoftSecond", allText);
                Assert.Contains("CellSoftFirst", allText);
                Assert.Contains("CellSoftSecond", allText);

                var words = pdf.GetPage(1).GetWords().ToList();
                double paragraphFirstY = Assert.Single(words, word => word.Text == "NativeSoftFirst").BoundingBox.Bottom;
                double paragraphSecondY = Assert.Single(words, word => word.Text == "NativeSoftSecond").BoundingBox.Bottom;
                double cellFirstY = Assert.Single(words, word => word.Text == "CellSoftFirst").BoundingBox.Bottom;
                double cellSecondY = Assert.Single(words, word => word.Text == "CellSoftSecond").BoundingBox.Bottom;

                Assert.True(paragraphFirstY > paragraphSecondY + 8D, $"Expected Word paragraph soft break to move following text to the next line. First y: {paragraphFirstY:0.##}, second y: {paragraphSecondY:0.##}.");
                Assert.True(cellFirstY > cellSecondY + 8D, $"Expected Word table cell soft break to move following text to the next line. First y: {cellFirstY:0.##}, second y: {cellSecondY:0.##}.");
                Assert.InRange(paragraphFirstY - paragraphSecondY, 10.5D, 14.5D);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Justified_Paragraphs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeJustifiedParagraph.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeJustifiedParagraph.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native justified paragraph alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau wraps across multiple visual lines.");
                paragraph.ParagraphAlignment = JustificationValues.Both;
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(240, 360),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(24)
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Native justified paragraph", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Matches(new Regex(@"(?:0\.[1-9]\d*|[1-9]\d*(?:\.\d+)?) Tw"), content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Shading_And_Uniform_Borders() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphPanel.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphPanel.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native shaded panel paragraph");
                paragraph.ShadingFillColorHex = "e6f2ff";
                paragraph.Borders.TopStyle = BorderValues.Single;
                paragraph.Borders.BottomStyle = BorderValues.Single;
                paragraph.Borders.LeftStyle = BorderValues.Single;
                paragraph.Borders.RightStyle = BorderValues.Single;
                paragraph.Borders.TopColorHex = "336699";
                paragraph.Borders.BottomColorHex = "336699";
                paragraph.Borders.LeftColorHex = "336699";
                paragraph.Borders.RightColorHex = "336699";
                paragraph.Borders.TopSize = 8;
                paragraph.Borders.BottomSize = 8;
                paragraph.Borders.LeftSize = 8;
                paragraph.Borders.RightSize = 8;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native shaded panel paragraph", text);
            }

            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("0.902 0.949 1 rg", raw);
            Assert.Contains("0.2 0.4 0.6 RG", raw);
            Assert.Contains("1 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Horizontal_Line() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHorizontalLine.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHorizontalLine.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native horizontal line");
                document.AddHorizontalLine(BorderValues.Single, OfficeIMO.Drawing.OfficeColor.Red, size: 16, space: 4);
                document.AddParagraph("After native horizontal line");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(300, 180),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Before native horizontal", text);
                Assert.Contains("After native horizontal", text);
            }

            Assert.Contains("1 0 0 RG", raw);
            Assert.Contains("2 w", raw);
            Assert.DoesNotContain("Before native horizontal lineAfter native horizontal line", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Bottom_Border_As_Rule() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBottomBorder.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBottomBorder.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native bordered paragraph heading");
                paragraph.Borders.BottomStyle = BorderValues.Single;
                paragraph.Borders.BottomColorHex = "336699";
                paragraph.Borders.BottomSize = 12;
                paragraph.Borders.BottomSpace = 3;
                paragraph.LineSpacingAfterPoints = 8;

                document.AddParagraph("After native bottom border");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Native bordered paragraph", text);
                Assert.Contains("After native bottom border", text);
            }

            Assert.Contains("0.2 0.4 0.6 RG", raw);
            Assert.Contains("1.5 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Top_Border_As_Rule() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTopBorder.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTopBorder.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native top border");

                WordParagraph paragraph = document.AddParagraph("Native top bordered paragraph");
                paragraph.Borders.TopStyle = BorderValues.Single;
                paragraph.Borders.TopColorHex = "008000";
                paragraph.Borders.TopSize = 16;
                paragraph.Borders.TopSpace = 4;
                paragraph.LineSpacingBeforePoints = 8;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Before native top border", text);
                Assert.Contains("Native top bordered", text);
                Assert.Contains("paragraph", text);
            }

            Assert.Contains("0 0.502 0 RG", raw);
            Assert.Contains("2 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_NonUniform_Paragraph_Borders_As_Panel_Sides() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphSideBorders.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphSideBorders.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native side bordered paragraph");
                paragraph.Borders.LeftStyle = BorderValues.Single;
                paragraph.Borders.LeftColorHex = "ff0000";
                paragraph.Borders.LeftSize = 12;
                paragraph.Borders.RightStyle = BorderValues.Single;
                paragraph.Borders.RightColorHex = "0000ff";
                paragraph.Borders.RightSize = 20;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Native side bordered", text);
                Assert.Contains("paragraph", text);
            }

            Assert.Contains("1 0 0 RG", raw);
            Assert.Contains("1.5 w", raw);
            Assert.Contains("0 0 1 RG", raw);
            Assert.Contains("2.5 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Tab_Leaders() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTabs.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTabs.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Revenue\t12");
                paragraph.AddTabStop(4320, TabStopValues.Right, TabStopLeaderCharValues.Dot);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 180),
                    Margins = PageMargins.Uniform(36)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                var page = pdf.GetPage(1);
                string text = page.Text;
                Assert.Contains("Revenue", text);
                Assert.Contains("12", text);

                int dotCount = page.Letters.Count(letter => letter.Value == ".");
                Assert.True(dotCount >= 15, $"Expected Word tab stop leaders to render across the native paragraph tab gap. Dot count: {dotCount}.");
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Pagination_And_Tab_Style() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyle.docx"));
            WordParagraph paragraph = document.AddParagraph("Native style flags");
            paragraph.KeepLinesTogether = true;
            paragraph.KeepWithNext = true;
            paragraph.AvoidWidowAndOrphan = true;
            paragraph.AddTabStop(1440, TabStopValues.Right, TabStopLeaderCharValues.Dot);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(72D, style.DefaultTabStopWidth);
            Assert.Equal(1.15D, style.LineHeight);
            Assert.Equal(8D, style.SpacingAfter);
            Assert.True(style.KeepTogether);
            Assert.True(style.KeepWithNext);
            Assert.True(style.WidowControl);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Explicit_Paragraph_Line_Spacing() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeExplicitParagraphStyle.docx"));
            WordParagraph exactParagraph = document.AddParagraph("Native exact line spacing");
            exactParagraph.FontSize = 12;
            exactParagraph.LineSpacingPoints = 24;

            WordParagraph autoParagraph = document.AddParagraph("Native auto line spacing");
            autoParagraph.LineSpacing = 276;
            autoParagraph.LineSpacingRule = LineSpacingRuleValues.Auto;

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
            PdfParagraphStyle exactStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { exactParagraph }));
            PdfParagraphStyle autoStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { autoParagraph }));

            Assert.Equal(2D, exactStyle.LineHeight);
            Assert.Equal(1.15D, autoStyle.LineHeight);
        }
    }
}
