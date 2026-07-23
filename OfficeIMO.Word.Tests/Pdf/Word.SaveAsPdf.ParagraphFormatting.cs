using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
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
                Assert.InRange(paragraphFirstY - paragraphSecondY, 10.5D, 15.5D);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Applies_Inherited_Paragraph_KeepWithNext() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeInheritedKeepWithNext.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeInheritedKeepWithNext.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "InheritedKeepWithNext";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Inherited Keep With Next" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new KeepNext()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph intro = document.AddParagraph("IntroMarker");
                intro.LineSpacingAfterPoints = 70;
                document.AddParagraph("InheritedKeepLabel").SetStyleId(styleId);
                document.AddParagraph("FollowingBody");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(260, 170),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);

            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
            Assert.DoesNotContain("InheritedKeepLabel", pdf.GetPage(1).Text);
            Assert.Contains("InheritedKeepLabel", pdf.GetPage(2).Text);
            Assert.Contains("FollowingBody", pdf.GetPage(2).Text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Keeps_Paragraph_KeepWithNext_Chains_Together() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeKeepWithNextChain.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeKeepWithNextChain.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "ChainKeepWithNext";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Chain Keep With Next" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new KeepNext()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph intro = document.AddParagraph("ChainIntro");
                intro.LineSpacingAfterPoints = 100;
                document.AddParagraph("ChainLead").SetStyleId(styleId);
                document.AddParagraph("ChainBridge").SetStyleId(styleId);
                document.AddParagraph("ChainTarget");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(260, 220),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);

            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("ChainIntro", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ChainLead", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ChainBridge", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ChainTarget", pdf.GetPage(1).Text);
            Assert.Contains("ChainLead", pdf.GetPage(2).Text);
            Assert.Contains("ChainBridge", pdf.GetPage(2).Text);
            Assert.Contains("ChainTarget", pdf.GetPage(2).Text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Applies_Inherited_PageBreakBefore() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeInheritedPageBreakBefore.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeInheritedPageBreakBefore.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "InheritedPageBreakBefore";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Inherited Page Break Before" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new PageBreakBefore()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                document.AddParagraph("BeforeInheritedPageBreak");
                document.AddParagraph("StyledPageBreakTarget").SetStyleId(styleId);
                document.AddParagraph("AfterStyledPageBreak");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);

            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("BeforeInheritedPageBreak", pdf.GetPage(1).Text);
            Assert.DoesNotContain("StyledPageBreakTarget", pdf.GetPage(1).Text);
            Assert.Contains("StyledPageBreakTarget", pdf.GetPage(2).Text);
            Assert.Contains("AfterStyledPageBreak", pdf.GetPage(2).Text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Text_After_Break_In_Same_Run() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBreakAndTextSameRun.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBreakAndTextSameRun.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document._document.Body!.Append(
                    new Paragraph(
                        new Run(new Text("BodyBefore")),
                        new Run(new Break(), new Text("BodyAfterSameRun"))));

                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0]._tableCell.RemoveAllChildren<Paragraph>();
                table.Rows[0].Cells[0]._tableCell.Append(
                    new Paragraph(
                        new Run(new Text("CellBefore")),
                        new Run(new Break(), new Text("CellAfterSameRun"))));

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("BodyBefore", allText);
            Assert.Contains("BodyAfterSameRun", allText);
            Assert.Contains("CellBefore", allText);
            Assert.Contains("CellAfterSameRun", allText);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Preserves_Empty_Paragraph_Line_Spacing() {
            double noBlankGap = RenderNativeEmptyParagraphGap("PdfNativeNoBlankParagraphSpacing", includeBlankParagraph: false);
            double blankGap = RenderNativeEmptyParagraphGap("PdfNativeBlankParagraphSpacing", includeBlankParagraph: true);

            Assert.True(blankGap > noBlankGap + 20D, $"Expected an empty Word paragraph to preserve line spacing in native PDF output. Gap without blank paragraph: {noBlankGap:0.##}; gap with blank paragraph: {blankGap:0.##}.");
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
                paragraph.ShadingFillColorHex = "E6F2FF";
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
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Style_Shading_And_Uniform_Borders() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphStylePanel.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphStylePanel.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeStylePanel";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Style Panel" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "E2F0D9" },
                        new ParagraphBorders(
                            new TopBorder { Val = BorderValues.Single, Color = "385723", Size = 8U },
                            new LeftBorder { Val = BorderValues.Single, Color = "385723", Size = 8U },
                            new BottomBorder { Val = BorderValues.Single, Color = "385723", Size = 8U },
                            new RightBorder { Val = BorderValues.Single, Color = "385723", Size = 8U })))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph paragraph = document.AddParagraph("Native styled panel paragraph");
                paragraph.SetStyleId(styleId);
                document.AddParagraph("After styled panel");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native styled panel paragraph", text);
                Assert.Contains("After styled panel", text);
            }

            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("0.886 0.941 0.851 rg", raw);
            Assert.Contains("0.22 0.341 0.137 RG", raw);
            Assert.Contains("1 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Joins_Adjacent_Borderless_Paragraph_Shading() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAdjacentParagraphShading.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAdjacentParagraphShading.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph first = document.AddParagraph("Adjacent shaded heading");
                first.ShadingFillColorHex = "F0E68C";
                WordParagraph second = document.AddParagraph("Adjacent shaded value");
                second.ShadingFillColorHex = "F0E68C";
                document.AddParagraph("After shaded block");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 240),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(36)
                });
            }

            string raw = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            var fills = ExtractFilledRectangles(raw, "0.941 0.902 0.549 rg")
                .Where(fill => fill.Width > 200D)
                .OrderByDescending(fill => fill.Y)
                .Take(2)
                .ToArray();

            Assert.Equal(2, fills.Length);
            double verticalGap = fills[0].Y - (fills[1].Y + fills[1].Height);
            Assert.InRange(Math.Abs(verticalGap), 0D, 0.25D);
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
                paragraph.Borders.LeftColorHex = "FF0000";
                paragraph.Borders.LeftSize = 12;
                paragraph.Borders.RightStyle = BorderValues.Single;
                paragraph.Borders.RightColorHex = "0000FF";
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
        public void SaveAsPdf_OfficeIMOEngine_Honors_Paragraph_Border_Space_As_Panel_Padding() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBorderSpace.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBorderSpace.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph tight = document.AddParagraph("TightSpace");
                tight.Borders.LeftStyle = BorderValues.Single;
                tight.Borders.LeftColorHex = "444444";
                tight.Borders.LeftSize = 8;
                tight.Borders.LeftSpace = 0;
                tight.LineSpacingAfterPoints = 4;

                WordParagraph wide = document.AddParagraph("WideSpace");
                wide.Borders.LeftStyle = BorderValues.Single;
                wide.Borders.LeftColorHex = "444444";
                wide.Borders.LeftSize = 8;
                wide.Borders.LeftSpace = 24;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 220),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            var tightWord = Assert.Single(words, word => word.Text == "TightSpace");
            var wideWord = Assert.Single(words, word => word.Text == "WideSpace");

            Assert.True(wideWord.BoundingBox.Left > tightWord.BoundingBox.Left + 18D,
                $"Expected Word paragraph border space to move text away from the border. Tight x: {tightWord.BoundingBox.Left:0.##}; wide x: {wideWord.BoundingBox.Left:0.##}.");
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
        public void SaveAsPdf_OfficeIMOEngine_Maps_Document_Default_Tab_Stop() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultTabStop.docx"));
            document.Settings.DefaultTabStop = 1440;
            WordParagraph paragraph = document.AddParagraph("Native document default tab stop");

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(72D, style.DefaultTabStopWidth);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Document_Default_Tab_Stop() {
            (double narrowLeftX, double narrowRightX) = RenderDocumentDefaultTabStop(720, "PdfNativeDocumentDefaultTabStopNarrow");
            (double wideLeftX, double wideRightX) = RenderDocumentDefaultTabStop(2160, "PdfNativeDocumentDefaultTabStopWide");

            Assert.InRange(Math.Abs(wideLeftX - narrowLeftX), 0D, 0.75D);
            Assert.True(wideRightX > narrowRightX + 50D,
                $"Expected wider Word document default tab stop to move implicit tab text right. Narrow x: {narrowRightX:0.##}, wide x: {wideRightX:0.##}.");
        }

        private (double LeftX, double RightX) RenderDocumentDefaultTabStop(int defaultTabStopTwips, string fileName) {
            string docPath = Path.Combine(_directoryWithFiles, fileName + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileName + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Settings.DefaultTabStop = defaultTabStopTwips;
                document.AddParagraph("WWWWWWWWWWWW\tTabRight");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(420, 180),
                    Margins = PageMargins.Uniform(36)
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var page = pdf.GetPage(1);
            Assert.Contains("TabRight", page.Text);
            return (FindWordStartX(page, "WWWWWWWWWWWW"), FindWordStartX(page, "TabRight"));
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Pagination_And_Tab_Style() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyle.docx"));
            WordParagraph paragraph = document.AddParagraph("Native style flags");
            paragraph.KeepLinesTogether = true;
            paragraph.KeepWithNext = true;
            paragraph.AvoidWidowAndOrphan = true;
            paragraph.AddTabStop(1440, TabStopValues.Right, TabStopLeaderCharValues.Dot);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Null(style.DefaultTabStopWidth);
            PdfTabStop tabStop = Assert.Single(style.TabStops);
            Assert.Equal(72D, tabStop.Position);
            Assert.Equal(PdfTabAlignment.Right, tabStop.Alignment);
            Assert.Equal(PdfTabLeaderStyle.Dots, tabStop.Leader);
            Assert.Equal(1.15D * (259D / 240D), style.LineHeight);
            Assert.Equal(8D, style.SpacingAfter);
            Assert.True(style.KeepTogether);
            Assert.True(style.KeepWithNext);
            Assert.True(style.WidowControl);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_TabStops() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleTabStops.docx"));
            const string styleId = "NativeStyleTabStops";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Style Tab Stops" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(
                    new Tabs(
                        new TabStop {
                            Val = TabStopValues.Right,
                            Leader = TabStopLeaderCharValues.Dot,
                            Position = 1440
                        })))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph paragraph = document.AddParagraph("Native style tab stops");
            paragraph.SetStyleId(styleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Null(style.DefaultTabStopWidth);
            PdfTabStop tabStop = Assert.Single(style.TabStops);
            Assert.Equal(72D, tabStop.Position);
            Assert.Equal(PdfTabAlignment.Right, tabStop.Alignment);
            Assert.Equal(PdfTabLeaderStyle.Dots, tabStop.Leader);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Style_Tab_Leaders() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleTabLeaders.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleTabLeaders.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeRenderedStyleTabStops";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Rendered Style Tab Stops" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(
                        new Tabs(
                            new TabStop {
                                Val = TabStopValues.Right,
                                Leader = TabStopLeaderCharValues.Dot,
                                Position = 4320
                            })))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph paragraph = document.AddParagraph("StyleRevenue\t42");
                paragraph.SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 180),
                    Margins = PageMargins.Uniform(36)
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var page = pdf.GetPage(1);
            Assert.Contains("StyleRevenue", page.Text);
            Assert.Contains("42", page.Text);

            int dotCount = page.Letters.Count(letter => letter.Value == ".");
            Assert.True(dotCount >= 15, $"Expected paragraph style tab leaders to render across the native paragraph tab gap. Dot count: {dotCount}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Applies_Paragraph_Style_Indentation() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleIndentation.docx"));
            const string styleId = "NativeStyleIndentation";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Style Indentation" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new Indentation {
                    Left = "1440",
                    Right = "720",
                    Hanging = "360"
                }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph paragraph = document.AddParagraph("Native style indentation");
            paragraph.SetStyleId(styleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(72D, style.LeftIndent);
            Assert.Equal(36D, style.RightIndent);
            Assert.Equal(-18D, style.FirstLineIndent);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_Alignment() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleAlignment.docx"));
            const string styleId = "NativeStyleAlignment";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Style Alignment" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new Justification {
                    Val = JustificationValues.Center
                }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph styled = document.AddParagraph("Native centered style alignment");
            styled.SetStyleId(styleId);
            WordParagraph directOverride = document.AddParagraph("Native direct alignment override");
            directOverride.SetStyleId(styleId);
            directOverride.ParagraphAlignment = JustificationValues.Right;

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("ResolveNativeParagraphAlign", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph), typeof(bool) }, modifiers: null)!;

            Assert.Equal(PdfAlign.Center, (PdfAlign)method.Invoke(null, new object[] { styled, true })!);
            Assert.Equal(PdfAlign.Right, (PdfAlign)method.Invoke(null, new object[] { directOverride, true })!);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Style_Alignment() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedParagraphStyleAlignment.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedParagraphStyleAlignment.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeRenderedStyleAlignment";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Rendered Style Alignment" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new Justification {
                        Val = JustificationValues.Center
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph styled = document.AddParagraph("StyledCenterMarker");
                styled.SetStyleId(styleId);
                document.AddParagraph("LeftMarker");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(300, 180),
                    Margins = PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            Assert.True(File.Exists(pdfPath));
            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            var centered = Assert.Single(words, word => word.Text == "StyledCenterMarker");
            var left = Assert.Single(words, word => word.Text == "LeftMarker");

            double centeredMidpoint = centered.BoundingBox.Left + (centered.BoundingBox.Width / 2D);
            Assert.InRange(centeredMidpoint, 130D, 170D);
            Assert.True(centered.BoundingBox.Left > left.BoundingBox.Left + 40D, $"Expected style-centered paragraph text to move right of left-aligned text. Centered x: {centered.BoundingBox.Left:0.##}; left x: {left.BoundingBox.Left:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_BiDi_Paragraphs_Right_Aligned() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBiDiParagraphAlignment.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBiDiParagraphAlignment.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph bidi = document.AddParagraph("BidiRightMarker");
                bidi.BiDi = true;
                document.AddParagraph("LeftMarker");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(300, 180),
                    Margins = PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
            var words = pdf.GetPage(1).GetWords().ToList();
            var bidiWord = Assert.Single(words, word => word.Text == "BidiRightMarker");
            var leftWord = Assert.Single(words, word => word.Text == "LeftMarker");

            Assert.True(bidiWord.BoundingBox.Left > leftWord.BoundingBox.Left + 90D, $"Expected BiDi paragraph text to use Word-style right alignment. BiDi x: {bidiWord.BoundingBox.Left:0.##}; left x: {leftWord.BoundingBox.Left:0.##}.");

            PdfDocumentInfo info = PdfInspector.Inspect(bytes);
            Assert.Equal("R2L", info.ViewerPreferences?.GetValue("Direction"));
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Preserves_Configured_Viewer_Direction_For_BiDi_Documents() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBiDiConfiguredViewerDirection.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBiDiConfiguredViewerDirection.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph bidi = document.AddParagraph("ConfiguredBidiMarker");
                bidi.BiDi = true;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PdfOptions = new PdfOptions {
                        ViewerPreferences = new PdfViewerPreferencesOptions {
                            Direction = PdfViewerDirection.LeftToRight
                        }
                    }
                });
            }

            PdfDocumentInfo info = PdfInspector.Inspect(File.ReadAllBytes(pdfPath));
            Assert.Equal("L2R", info.ViewerPreferences?.GetValue("Direction"));
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_Run_Formatting() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleRunFormatting.docx"));
            const string styleId = "NativeStyleRunFormatting";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Style Run Formatting" },
                new BasedOn { Val = "Normal" },
                new StyleRunProperties(
                    new Bold(),
                    new Italic(),
                    new Underline { Val = UnderlineValues.Single },
                    new Color { Val = "C00000" },
                    new Highlight { Val = HighlightColorValues.Yellow },
                    new FontSize { Val = "28" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph paragraph = document.AddParagraph("Native style run formatting");
            paragraph.SetStyleId(styleId);
            WordParagraph directOverride = document.AddParagraph("Native direct run override");
            directOverride.SetStyleId(styleId);
            directOverride._run!.RunProperties ??= new RunProperties();
            directOverride._run.RunProperties.Bold = new Bold { Val = false };
            directOverride._run.RunProperties.Italic = new Italic { Val = false };
            directOverride._run.RunProperties.Underline = new Underline { Val = UnderlineValues.None };
            directOverride._run.RunProperties.Color = new Color { Val = "0000FF" };
            directOverride._run.RunProperties.FontSize = new FontSize { Val = "20" };

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeCellParagraphRuns", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph), typeof(Dictionary<long, int>) }, modifiers: null)!;
            var runs = Assert.IsAssignableFrom<IReadOnlyList<TextRun>>(method.Invoke(null, new object?[] { paragraph, null }));
            TextRun run = Assert.Single(runs);
            var overrideRuns = Assert.IsAssignableFrom<IReadOnlyList<TextRun>>(method.Invoke(null, new object?[] { directOverride, null }));
            TextRun overrideRun = Assert.Single(overrideRuns);

            Assert.True(run.Bold);
            Assert.True(run.Italic);
            Assert.True(run.Underline);
            Assert.Equal(PdfColor.FromRgb(192, 0, 0), run.Color);
            Assert.Equal(PdfColor.FromRgb(255, 255, 0), run.BackgroundColor);
            Assert.Equal(14D, run.FontSize);
            Assert.False(overrideRun.Bold);
            Assert.False(overrideRun.Italic);
            Assert.False(overrideRun.Underline);
            Assert.Equal(PdfColor.FromRgb(0, 0, 255), overrideRun.Color);
            Assert.Equal(10D, overrideRun.FontSize);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Style_Run_Formatting() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedParagraphStyleRunFormatting.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedParagraphStyleRunFormatting.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeRenderedStyleRunFormatting";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Rendered Style Run Formatting" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(
                        new Color { Val = "C00000" },
                        new FontSize { Val = "28" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph styled = document.AddParagraph("StyledRunFormatMarker");
                styled.SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                Assert.Contains("StyledRunFormatMarker", pdf.GetPage(1).Text);
            }

            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("0.753 0 0 rg", raw);
            Assert.Matches(new Regex(@"/F\d+ 14 Tf"), raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Omits_Hidden_Body_Text_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHiddenBodyText.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHiddenBodyText.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string hiddenStyleId = "NativeHiddenRunStyle";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Hidden Run Style" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(new Vanish()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = hiddenStyleId,
                    CustomStyle = true
                });

                WordParagraph mixed = document.AddParagraph();
                mixed.AddText("VisibleStart");
                WordParagraph hiddenRun = mixed.AddText("HiddenBodyRun");
                hiddenRun._run!.RunProperties ??= new RunProperties();
                hiddenRun._run.RunProperties.Vanish = new Vanish();
                mixed.AddText("VisibleEnd");

                WordParagraph hiddenOnly = document.AddParagraph("HiddenOnlyParagraph");
                hiddenOnly._run!.RunProperties ??= new RunProperties();
                hiddenOnly._run.RunProperties.Vanish = new Vanish();

                WordParagraph hiddenByStyle = document.AddParagraph("HiddenByStyle");
                hiddenByStyle.SetStyleId(hiddenStyleId);

                document.AddParagraph("VisibleAfterHiddenText");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("VisibleStart", text);
            Assert.Contains("VisibleEnd", text);
            Assert.Contains("VisibleAfterHiddenText", text);
            Assert.DoesNotContain("HiddenBodyRun", text);
            Assert.DoesNotContain("HiddenOnlyParagraph", text);
            Assert.DoesNotContain("HiddenByStyle", text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Caps_Body_Text_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCapsBodyText.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCapsBodyText.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string capsStyleId = "NativeCapsRunStyle";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Caps Run Style" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(new Caps()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = capsStyleId,
                    CustomStyle = true
                });

                WordParagraph direct = document.AddParagraph();
                direct.AddText("beforeCaps ");
                WordParagraph capsRun = direct.AddText("capsBodyRun");
                capsRun._run!.RunProperties ??= new RunProperties();
                capsRun._run.RunProperties.Caps = new Caps();
                direct.AddText(" afterCaps");

                WordParagraph styled = document.AddParagraph("capsStyleRun");
                styled.SetStyleId(capsStyleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("beforeCaps", text);
            Assert.Contains("CAPSBODYRUN", text);
            Assert.Contains("CAPSSTYLERUN", text);
            Assert.DoesNotContain("capsBodyRun", text);
            Assert.DoesNotContain("capsStyleRun", text);
        }

        private static IReadOnlyList<(double X, double Y, double Width, double Height)> ExtractFilledRectangles(string rawPdf, string colorOperator) {
            string pattern = Regex.Escape(colorOperator) +
                @"\s+(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) (?<width>-?\d+(?:\.\d+)?) (?<height>-?\d+(?:\.\d+)?) re f";
            return Regex.Matches(rawPdf, pattern)
                .Cast<Match>()
                .Select(match => (
                    X: double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                    Y: double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                    Width: double.Parse(match.Groups["width"].Value, CultureInfo.InvariantCulture),
                    Height: double.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture)))
                .ToArray();
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Defaults_Paragraph_WidowControl_To_Word_Semantics() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDefaultWidowControl.docx"));
            WordParagraph paragraph = document.AddParagraph("Native default widow control");

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.True(style.WidowControl);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_Explicit_Paragraph_WidowControl_Off() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeExplicitWidowControlOff.docx"));
            WordParagraph paragraph = document.AddParagraph("Native explicit widow control off");
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.WidowControl = new WidowControl {
                Val = false
            };

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.False(style.WidowControl);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Uses_Document_Default_Run_Font_Family() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultRunFont.docx"));
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.DocDefaults ??= new DocDefaults();
            RunPropertiesDefault runDefaults = styles.DocDefaults.GetFirstChild<RunPropertiesDefault>() ?? styles.DocDefaults.AppendChild(new RunPropertiesDefault());
            RunPropertiesBaseStyle runProperties = runDefaults.GetFirstChild<RunPropertiesBaseStyle>() ?? runDefaults.AppendChild(new RunPropertiesBaseStyle());
            runProperties.RunFonts = new RunFonts {
                Ascii = "Times New Roman",
                HighAnsi = "Times New Roman"
            };

            const string styleId = "NativeNoRunFontStyle";
            styles.Append(new Style(new StyleName { Val = "Native No Run Font Style" }) {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph paragraph = document.AddParagraph("Native document default font");
            paragraph.SetStyleId(styleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeCellParagraphRuns", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph), typeof(Dictionary<long, int>) }, modifiers: null)!;
            var runs = Assert.IsAssignableFrom<IReadOnlyList<TextRun>>(method.Invoke(null, new object?[] { paragraph, null }));
            TextRun run = Assert.Single(runs);

            Assert.Equal(PdfStandardFont.TimesRoman, run.Font);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Reports_Document_Default_Font_Substitution() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultFontSubstitution.docx"));
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.DocDefaults ??= new DocDefaults();
            RunPropertiesDefault runDefaults = styles.DocDefaults.GetFirstChild<RunPropertiesDefault>() ?? styles.DocDefaults.AppendChild(new RunPropertiesDefault());
            RunPropertiesBaseStyle runProperties = runDefaults.GetFirstChild<RunPropertiesBaseStyle>() ?? runDefaults.AppendChild(new RunPropertiesBaseStyle());
            runProperties.RunFonts = new RunFonts {
                Ascii = "Calibri",
                HighAnsi = "Calibri"
            };
            document.AddParagraph("Document default font diagnostic");

            PdfDocumentConversionResult result = document.ToPdfDocumentResult(new PdfSaveOptions {
                IncludePageNumbers = false,
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
            });
            _ = result.ToBytes();

            PdfConversionWarning warning = Assert.Single(result.Warnings, item =>
                item.Code == "NativeFontFamilySubstituted" &&
                item.Details.TryGetValue("fontFamily", out string? family) &&
                family == "Calibri");
            Assert.Equal("Helvetica", warning.Details["fallbackSlot"]);
            Assert.False(warning.Details.ContainsKey("resolvedFontFamily"));
            Assert.DoesNotContain("embedded family", warning.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Document_Default_Language_To_Pdf_Catalog() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDocumentDefaultLanguage.docx"));
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.DocDefaults ??= new DocDefaults();
            RunPropertiesDefault runDefaults = styles.DocDefaults.GetFirstChild<RunPropertiesDefault>() ?? styles.DocDefaults.AppendChild(new RunPropertiesDefault());
            RunPropertiesBaseStyle runProperties = runDefaults.GetFirstChild<RunPropertiesBaseStyle>() ?? runDefaults.AppendChild(new RunPropertiesBaseStyle());
            runProperties.Languages = new Languages { Val = "pl-PL" };
            document.AddParagraph("Document language");

            byte[] bytes = document.ToPdf(new PdfSaveOptions {
                IncludePageNumbers = false
            });

            Assert.Equal("pl-PL", PdfReadDocument.Open(bytes).CatalogLanguage);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Ignores_Bar_And_Clear_TabStops_For_Text_Tabs() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeIgnoredTabStops.docx"));
            WordParagraph paragraph = document.AddParagraph("Native ignored tab stops");
            paragraph.AddTabStop(720, TabStopValues.Bar, TabStopLeaderCharValues.None);
            paragraph.AddTabStop(1440, TabStopValues.Clear, TabStopLeaderCharValues.None);
            paragraph.AddTabStop(2160, TabStopValues.Right, TabStopLeaderCharValues.Dot);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            PdfTabStop tabStop = Assert.Single(style.TabStops);
            Assert.Null(style.DefaultTabStopWidth);
            Assert.Equal(108D, tabStop.Position);
            Assert.Equal(PdfTabAlignment.Right, tabStop.Alignment);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Keeps_Default_Tab_Width_Separate_From_Explicit_TabStops() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDefaultTabWidth.docx"));
            WordParagraph paragraph = document.AddParagraph("Native default tab width");
            paragraph.AddTabStop(2880, TabStopValues.Left, TabStopLeaderCharValues.None);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Null(style.DefaultTabStopWidth);
            PdfTabStop tabStop = Assert.Single(style.TabStops);
            Assert.Equal(144D, tabStop.Position);
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

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle exactStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { exactParagraph }));
            PdfParagraphStyle autoStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { autoParagraph }));

            Assert.Equal(2D, exactStyle.LineHeight);
            Assert.Equal(1.15D * (276D / 240D), autoStyle.LineHeight);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_AtLeast_Paragraph_Line_Spacing() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeAtLeastParagraphStyle.docx"));
            WordParagraph directParagraph = document.AddParagraph("Native at least line spacing");
            directParagraph.FontSize = 24;
            directParagraph.LineSpacingPoints = 6;
            directParagraph.LineSpacingRule = LineSpacingRuleValues.AtLeast;

            const string styleId = "AtLeastLineSpacingStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "AtLeast Line Spacing Style" },
                new StyleRunProperties(new FontSize { Val = "48" }),
                new StyleParagraphProperties(new SpacingBetweenLines {
                    Line = "120",
                    LineRule = LineSpacingRuleValues.AtLeast
                }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph styledParagraph = document.AddParagraph("Native style at least line spacing");
            styledParagraph.SetStyleId(styleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle directStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { directParagraph }));
            PdfParagraphStyle inheritedStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { styledParagraph }));

            Assert.NotNull(directStyle.LineHeight);
            Assert.NotNull(inheritedStyle.LineHeight);
            Assert.InRange(directStyle.LineHeight.Value, 1.2D, 1.3D);
            Assert.InRange(inheritedStyle.LineHeight.Value, 1.2D, 1.3D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_AtLeast_Paragraph_Line_Spacing() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedAtLeastLineSpacing.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedAtLeastLineSpacing.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph exact = document.AddParagraph("ExactSmallFirst");
                exact.FontSize = 24;
                exact.LineSpacingPoints = 6;
                exact.LineSpacingRule = LineSpacingRuleValues.Exact;
                exact.AddBreak();
                exact.AddText("ExactSmallSecond");

                WordParagraph atLeast = document.AddParagraph("AtLeastFirst");
                atLeast.FontSize = 24;
                atLeast.LineSpacingPoints = 6;
                atLeast.LineSpacingRule = LineSpacingRuleValues.AtLeast;
                atLeast.AddBreak();
                atLeast.AddText("AtLeastSecond");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 260),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double exactFirstY = Assert.Single(words, word => word.Text == "ExactSmallFirst").BoundingBox.Bottom;
            double exactSecondY = Assert.Single(words, word => word.Text == "ExactSmallSecond").BoundingBox.Bottom;
            double atLeastFirstY = Assert.Single(words, word => word.Text == "AtLeastFirst").BoundingBox.Bottom;
            double atLeastSecondY = Assert.Single(words, word => word.Text == "AtLeastSecond").BoundingBox.Bottom;
            double exactGap = exactFirstY - exactSecondY;
            double atLeastGap = atLeastFirstY - atLeastSecondY;

            Assert.True(atLeastGap > exactGap + 14D, $"Expected Word atLeast line spacing to preserve natural line advance instead of exact compressed leading. Exact gap: {exactGap:0.##}; atLeast gap: {atLeastGap:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Uses_Character_Style_Font_Size_For_Exact_Line_Spacing() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCharacterStyleExactLineSpacing.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCharacterStyleExactLineSpacing.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeExactLineCharacterStyle";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native Exact Line Character Style" },
                    new StyleRunProperties(new FontSize { Val = "64" }))
                {
                    Type = StyleValues.Character,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("CharExactFirst").SetCharacterStyleId(styleId);
                paragraph.AddBreak();
                paragraph.AddText("CharExactSecond").SetCharacterStyleId(styleId);
                paragraph.LineSpacingAfterPoints = 0;
                paragraph.LineSpacingPoints = 18;
                paragraph.LineSpacingRule = LineSpacingRuleValues.Exact;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double firstY = Assert.Single(words, word => word.Text == "CharExactFirst").BoundingBox.Bottom;
            double secondY = Assert.Single(words, word => word.Text == "CharExactSecond").BoundingBox.Bottom;

            Assert.InRange(firstY - secondY, 16D, 22D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_Exact_Line_Spacing() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleExactLineSpacing.docx"));
            const string styleId = "ExactLineSpacingStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Exact Line Spacing Style" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new SpacingBetweenLines {
                    Line = "480",
                    LineRule = LineSpacingRuleValues.Exact
                }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph paragraph = document.AddParagraph("Native style exact line spacing");
            paragraph.FontSize = 12;
            paragraph.SetStyleId(styleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(2D, style.LineHeight);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_Line_Unit_Spacing() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleLineUnitSpacing.docx"));
            const string styleId = "LineUnitSpacingStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Line Unit Spacing Style" },
                new BasedOn { Val = "Normal" },
                new StyleRunProperties(new FontSize { Val = "24" }),
                new StyleParagraphProperties(new SpacingBetweenLines {
                    BeforeLines = 50,
                    AfterLines = 150
                }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph paragraph = document.AddParagraph("Native style line unit spacing");
            paragraph.SetStyleId(styleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(6.9D, style.SpacingBefore, 3);
            Assert.Equal(20.7D, style.SpacingAfter!.Value, 3);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Style_Line_Unit_Spacing() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedParagraphLineUnitSpacing.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRenderedParagraphLineUnitSpacing.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "RenderedLineUnitSpacingStyle";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Rendered Line Unit Spacing Style" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(new FontSize { Val = "24" }),
                    new StyleParagraphProperties(new SpacingBetweenLines {
                        BeforeLines = 100,
                        AfterLines = 150
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph before = document.AddParagraph("LineUnitBefore");
                before.LineSpacingAfterPoints = 0;
                WordParagraph styled = document.AddParagraph("LineUnitStyled");
                styled.SetStyleId(styleId);
                WordParagraph after = document.AddParagraph("LineUnitAfter");
                after.LineSpacingAfterPoints = 0;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(320, 240),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double beforeY = Assert.Single(words, word => word.Text == "LineUnitBefore").BoundingBox.Bottom;
            double styledY = Assert.Single(words, word => word.Text == "LineUnitStyled").BoundingBox.Bottom;
            double afterY = Assert.Single(words, word => word.Text == "LineUnitAfter").BoundingBox.Bottom;

            Assert.True(beforeY > styledY + 22D, $"Expected beforeLines spacing to push the styled paragraph down. Before y: {beforeY:0.##}; styled y: {styledY:0.##}.");
            Assert.True(styledY > afterY + 30D, $"Expected afterLines spacing to push the following paragraph down. Styled y: {styledY:0.##}; after y: {afterY:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Collapses_Adjacent_Paragraph_Spacing() {
            double afterOnlyGap = RenderNativeAdjacentParagraphSpacingGap("PdfNativeCollapsedParagraphSpacingAfterOnly", secondSpacingBefore: 0D);
            double collapsedGap = RenderNativeAdjacentParagraphSpacingGap("PdfNativeCollapsedParagraphSpacingBeforeSmallerThanAfter", secondSpacingBefore: 20D);

            Assert.InRange(Math.Abs(collapsedGap - afterOnlyGap), 0D, 2D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_Paragraph_Style_Contextual_Spacing() {
            double sameStyleGap = RenderNativeContextualSpacingGap("PdfNativeContextualSpacingSameStyle", sameStyle: true);
            double differentStyleGap = RenderNativeContextualSpacingGap("PdfNativeContextualSpacingDifferentStyle", sameStyle: false);

            Assert.True(differentStyleGap > sameStyleGap + 16D, $"Expected Word contextual spacing to suppress spacing between paragraphs with the same style. Same-style gap: {sameStyleGap:0.##}; different-style gap: {differentStyleGap:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Lets_Derived_Paragraph_Style_Auto_Line_Spacing_Override_Exact() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyleAutoOverridesExactLineSpacing.docx"));
            const string baseStyleId = "BaseExactLineSpacingStyle";
            const string derivedStyleId = "DerivedAutoLineSpacingStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Base Exact Line Spacing Style" },
                    new StyleParagraphProperties(new SpacingBetweenLines {
                        Line = "480",
                        LineRule = LineSpacingRuleValues.Exact
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = baseStyleId,
                    CustomStyle = true
                },
                new Style(
                    new StyleName { Val = "Derived Auto Line Spacing Style" },
                    new BasedOn { Val = baseStyleId },
                    new StyleParagraphProperties(new SpacingBetweenLines {
                        Line = "240",
                        LineRule = LineSpacingRuleValues.Auto
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = derivedStyleId,
                    CustomStyle = true
                });

            WordParagraph paragraph = document.AddParagraph("Native derived style auto line spacing");
            paragraph.FontSize = 12;
            paragraph.SetStyleId(derivedStyleId);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(1.15D, style.LineHeight);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Hanging_Indent() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeHangingIndentStyle.docx"));
            WordParagraph paragraph = document.AddParagraph("Native hanging paragraph");
            paragraph.IndentationBeforePoints = 72;
            paragraph.IndentationHangingPoints = 36;

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static, binder: null, new[] { typeof(WordParagraph) }, modifiers: null)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(72D, style.LeftIndent);
            Assert.Equal(-36D, style.FirstLineIndent);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Hanging_Indent_Without_Left_Indent() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHangingIndentNoLeft.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHangingIndentNoLeft.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("HangingOnly alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu");
                paragraph.IndentationHangingPoints = 36;

                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));
            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            Assert.Contains("HangingOnly", pdf.GetPage(1).Text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Hanging_Indent() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHangingIndent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHangingIndent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("HangingStart alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu");
                paragraph.IndentationBeforePoints = 72;
                paragraph.IndentationHangingPoints = 36;
                paragraph.LineSpacingAfterPoints = 0;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(260, 260),
                    Margins = PageMargins.Uniform(36)
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var lineLefts = pdf.GetPage(1)
                .GetWords()
                .GroupBy(word => Math.Round(word.BoundingBox.Bottom, 1))
                .OrderByDescending(group => group.Key)
                .Take(2)
                .Select(group => group.Min(word => word.BoundingBox.Left))
                .ToList();

            Assert.Equal(2, lineLefts.Count);
            Assert.True(lineLefts[1] > lineLefts[0] + 20D, $"Expected wrapped hanging-indent line to start farther right. First line x: {lineLefts[0]:0.##}; second line x: {lineLefts[1]:0.##}.");
        }

        private double RenderNativeEmptyParagraphGap(string fileNamePrefix, bool includeBlankParagraph) {
            string beforeMarker = fileNamePrefix + "Before";
            string afterMarker = fileNamePrefix + "After";
            string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph before = document.AddParagraph(beforeMarker);
                before.LineSpacingAfterPoints = 0;

                if (includeBlankParagraph) {
                    WordParagraph blank = document.AddParagraph();
                    blank.FontSize = 20;
                    blank.LineSpacingBeforePoints = 4;
                    blank.LineSpacingAfterPoints = 0;
                }

                WordParagraph after = document.AddParagraph(afterMarker);
                after.LineSpacingAfterPoints = 0;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(320, 240),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double beforeY = Assert.Single(words, word => word.Text == beforeMarker).BoundingBox.Bottom;
            double afterY = Assert.Single(words, word => word.Text == afterMarker).BoundingBox.Bottom;
            return beforeY - afterY;
        }

        private double RenderNativeAdjacentParagraphSpacingGap(string fileNamePrefix, double secondSpacingBefore) {
            const string firstMarker = "CollapseFirst";
            const string secondMarker = "CollapseSecond";
            string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph first = document.AddParagraph(firstMarker);
                first.LineSpacingAfterPoints = 30;
                WordParagraph second = document.AddParagraph(secondMarker);
                second.LineSpacingBeforePoints = secondSpacingBefore;
                second.LineSpacingAfterPoints = 0;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(320, 240),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double firstY = Assert.Single(words, word => word.Text == firstMarker).BoundingBox.Bottom;
            double secondY = Assert.Single(words, word => word.Text == secondMarker).BoundingBox.Bottom;
            return firstY - secondY;
        }

        private double RenderNativeContextualSpacingGap(string fileNamePrefix, bool sameStyle) {
            const string firstMarker = "ContextFirst";
            const string secondMarker = "ContextSecond";
            string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string contextualStyleId = "NativeContextualSpacing";
                const string otherStyleId = "NativeContextualSpacingOther";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(
                    new Style(
                        new StyleName { Val = "Native Contextual Spacing" },
                        new StyleParagraphProperties(
                            new SpacingBetweenLines { After = "480" },
                            new ContextualSpacing()))
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = contextualStyleId,
                        CustomStyle = true
                    },
                    new Style(
                        new StyleName { Val = "Native Contextual Spacing Other" },
                        new StyleParagraphProperties(new SpacingBetweenLines { After = "0" }))
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = otherStyleId,
                        CustomStyle = true
                    });

                WordParagraph first = document.AddParagraph(firstMarker);
                first.SetStyleId(contextualStyleId);
                WordParagraph second = document.AddParagraph(secondMarker);
                second.SetStyleId(sameStyle ? contextualStyleId : otherStyleId);
                second.LineSpacingAfterPoints = 0;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(320, 240),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double firstY = Assert.Single(words, word => word.Text == firstMarker).BoundingBox.Bottom;
            double secondY = Assert.Single(words, word => word.Text == secondMarker).BoundingBox.Bottom;
            return firstY - secondY;
        }

    }
}
