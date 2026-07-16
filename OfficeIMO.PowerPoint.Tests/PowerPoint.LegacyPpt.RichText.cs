using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string RichTextFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "RichTextPowerPoint.ppt");

        [Fact]
        public void NeutralReader_DecodesMicrosoftCharacterRuns() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(RichTextFixturePath);
            LegacyPptShape shape = Assert.Single(Assert.Single(legacy.Slides).Shapes,
                item => item.Text.Contains("Bold red", StringComparison.Ordinal));
            LegacyPptTextBody textBody = shape.TextBody;

            Assert.False(textBody.IsStyleTruncated);
            Assert.Equal("Bold red | Italic green | Underlined blue\nRegular second paragraph",
                textBody.Text);
            Assert.True(textBody.HasParagraphFormatting);
            Assert.False(textBody.HasUnprojectedParagraphFormatting);
            LegacyPptParagraphRun paragraphRun = Assert.Single(textBody.ParagraphRuns);
            Assert.Equal(0, paragraphRun.Start);
            Assert.Equal(textBody.Text.Length, paragraphRun.Length);
            Assert.Equal((ushort)0, paragraphRun.IndentLevel);
            Assert.True(paragraphRun.CharacterWrap);
            Assert.Null(paragraphRun.WordWrap);
            Assert.True(paragraphRun.Overflow);
            Assert.False(paragraphRun.HasUnprojectedFormatting);
            Assert.False(textBody.HasUnprojectedCharacterFormatting);
            LegacyPptFont arial = Assert.Single(legacy.Fonts,
                font => font.Index == 1 && font.Typeface == "Arial");
            Assert.True(arial.IsTrueType);
            Assert.False(arial.HasEmbeddedData);
            Assert.Equal(legacy.Fonts.Count, legacy.CreateImportReport().FontCount);
            Assert.Equal(0, legacy.CreateImportReport().EmbeddedFontCount);
            Assert.Equal(0, legacy.CreateImportReport().TextRulerCount);
            Assert.Equal(8, legacy.CreateImportReport().MasterTextStyleCount);
            Assert.Equal(32, legacy.CreateImportReport().MasterTextStyleLevelCount);
            Assert.Collection(textBody.CharacterRuns,
                run => AssertRun(run, "Bold red", 0, bold: true, size: 32, color: "C00000"),
                run => AssertRun(run, " | ", 8, size: 24, color: "222222"),
                run => AssertRun(run, "Italic green", 11, italic: true, size: 26, color: "008000"),
                run => AssertRun(run, " | ", 23, size: 24, color: "222222"),
                run => AssertRun(run, "Underlined blue\n", 26, underline: true, size: 22,
                    color: "0000C0"),
                run => AssertRun(run, "Regular second paragraph", 42, size: 20, color: "333333"));
            Assert.All(textBody.CharacterRuns, run => {
                Assert.Equal((ushort)1, run.FontIndex);
                Assert.Equal("Arial", run.Typeface);
                Assert.False(run.HasUnprojectedFormatting);
            });
            Assert.DoesNotContain(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-PARAGRAPH-PARTIAL");
            Assert.DoesNotContain(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-CHARACTER-PARTIAL");
            LegacyPptMaster master = Assert.Single(legacy.Masters);
            Assert.Equal(8, master.TextMasterStyles.Count);
            Assert.All(master.TextMasterStyles, style => Assert.False(style.IsTruncated));
            LegacyPptTextMasterStyle titleStyle = Assert.Single(master.TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Title);
            LegacyPptTextMasterStyleLevel titleLevel = Assert.Single(titleStyle.Levels);
            Assert.Equal((short)44, titleLevel.CharacterProperties.FontSizePoints);
            Assert.Equal("Calibri Light", titleLevel.CharacterProperties.Typeface);
            Assert.DoesNotContain(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-MASTER-STYLE-PRESERVE-ONLY"
                    || diagnostic.Code == "PPT-TEXT-MASTER-STYLE-TRUNCATED"
                    || diagnostic.Code == "PPT-TEXT-MASTER-STYLE-PARTIAL");
            Assert.DoesNotContain(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-STYLE-TRUNCATED");
        }

        [Fact]
        public void NormalLoad_ProjectsMicrosoftCharacterRunsAndPreservesBinaryExactly() {
            byte[] source = File.ReadAllBytes(RichTextFixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(RichTextFixturePath);
            P.Shape shape = Assert.IsType<P.Shape>(Assert.Single(
                Assert.Single(presentation.Slides).TextBoxes).Element);
            A.Paragraph[] paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToArray();
            P.SlideMaster slideMaster = presentation.Slides[0].SlidePart.SlideLayoutPart!
                .SlideMasterPart!.SlideMaster!;
            A.Level1ParagraphProperties nativeTitle = Assert.IsType<A.Level1ParagraphProperties>(
                slideMaster.TextStyles!.TitleStyle!.FirstChild);
            Assert.Equal(0, nativeTitle.LeftMargin!.Value);
            Assert.Equal(0, nativeTitle.Indent!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Left, nativeTitle.Alignment!.Value);
            A.DefaultRunProperties nativeTitleRun = nativeTitle.GetFirstChild<A.DefaultRunProperties>()!;
            Assert.Equal(4400, nativeTitleRun.FontSize!.Value);
            Assert.Equal("Calibri Light", nativeTitleRun.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            A.Level1ParagraphProperties nativeBody = Assert.IsType<A.Level1ParagraphProperties>(
                slideMaster.TextStyles.BodyStyle!.FirstChild);
            Assert.Equal(228600, nativeBody.LeftMargin!.Value);
            Assert.Equal(0, nativeBody.Indent!.Value);
            Assert.Equal(2800, nativeBody.GetFirstChild<A.DefaultRunProperties>()!.FontSize!.Value);

            Assert.Equal(2, paragraphs.Length);
            A.Run[] firstParagraph = paragraphs[0].Elements<A.Run>().ToArray();
            Assert.Equal(5, firstParagraph.Length);
            Assert.True(paragraphs[0].ParagraphProperties!.EastAsianLineBreak!.Value);
            Assert.True(paragraphs[0].ParagraphProperties.Height!.Value);
            Assert.True(paragraphs[1].ParagraphProperties!.EastAsianLineBreak!.Value);
            Assert.True(paragraphs[1].ParagraphProperties.Height!.Value);
            Assert.Equal("Bold red", firstParagraph[0].Text!.Text);
            Assert.True(firstParagraph[0].RunProperties!.Bold!.Value);
            Assert.Equal(3200, firstParagraph[0].RunProperties.FontSize!.Value);
            Assert.Equal("C00000", GetRunColor(firstParagraph[0]));
            Assert.Equal("Arial", firstParagraph[0].RunProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.True(firstParagraph[2].RunProperties!.Italic!.Value);
            Assert.Equal(2600, firstParagraph[2].RunProperties.FontSize!.Value);
            Assert.Equal("008000", GetRunColor(firstParagraph[2]));
            Assert.Equal("Arial", firstParagraph[2].RunProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal(A.TextUnderlineValues.Single,
                firstParagraph[4].RunProperties!.Underline!.Value);
            Assert.Equal(2200, firstParagraph[4].RunProperties.FontSize!.Value);
            Assert.Equal("0000C0", GetRunColor(firstParagraph[4]));
            Assert.Equal("Arial", firstParagraph[4].RunProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            A.Run secondParagraph = Assert.Single(paragraphs[1].Elements<A.Run>());
            Assert.Equal("Regular second paragraph", secondParagraph.Text!.Text);
            Assert.Equal(2000, secondParagraph.RunProperties!.FontSize!.Value);
            Assert.Equal("333333", GetRunColor(secondParagraph));
            Assert.Equal("Arial", secondParagraph.RunProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(source, presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_AuthorsShapeRichTextParagraphsBulletsAndFonts() {
            byte[] bytes;
            using (PowerPointPresentation presentation = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 360, 180);
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                var firstProperties = new A.ParagraphProperties(
                    new A.LineSpacing(
                        new A.SpacingPercent { Val = 125000 }),
                    new A.SpaceBefore(
                        new A.SpacingPoints { Val = 1250 }),
                    new A.BulletColor(
                        new A.RgbColorModelHex { Val = "123456" }),
                    new A.BulletSizePercentage { Val = 120000 },
                    new A.BulletFont { Typeface = "OfficeIMO Bullet" },
                    new A.CharacterBullet { Char = "•" },
                    new A.TabStopList(new A.TabStop {
                        Position = 476250,
                        Alignment = A.TextTabAlignmentValues.Decimal
                    })) {
                    Level = 1,
                    Alignment = A.TextAlignmentTypeValues.Center,
                    LeftMargin = 317500,
                    Indent = -158750,
                    DefaultTabSize = 635000,
                    FontAlignment = A.TextFontAlignmentValues.Center,
                    RightToLeft = true,
                    EastAsianLineBreak = true,
                    LatinLineBreak = false,
                    Height = true
                };
                var bold = new A.Run(
                    new A.RunProperties(
                        new A.SolidFill(
                            new A.RgbColorModelHex { Val = "C00000" }),
                        new A.LatinFont { Typeface = "OfficeIMO Latin" }) {
                        Bold = true,
                        FontSize = 3200,
                        Baseline = 10000
                    },
                    new A.Text("Bold red"));
                var italic = new A.Run(
                    new A.RunProperties(
                        new A.SolidFill(
                            new A.SchemeColor {
                                Val = A.SchemeColorValues.Accent1
                            }),
                        new A.LatinFont { Typeface = "OfficeIMO Latin" }) {
                        Italic = true,
                        FontSize = 2400
                    },
                    new A.Text(" italic"));
                var first = new A.Paragraph(firstProperties, bold, italic,
                    new A.EndParagraphRunProperties { FontSize = 2000 });
                var second = new A.Paragraph(
                    new A.ParagraphProperties {
                        Alignment = A.TextAlignmentTypeValues.Right
                    },
                    new A.Run(
                        new A.RunProperties(
                            new A.SolidFill(
                                new A.RgbColorModelHex { Val = "333333" }),
                            new A.LatinFont {
                                Typeface = "OfficeIMO Latin"
                            }) {
                            Underline = A.TextUnderlineValues.Single,
                            FontSize = 2000
                        },
                        new A.Text("Second paragraph")));
                shape.TextBody = new P.TextBody(new A.BodyProperties(),
                    new A.ListStyle(), first, second);

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptShape binaryShape = Assert.Single(
                Assert.Single(binary.Slides).Shapes,
                candidate => candidate.Text.StartsWith("Bold red",
                    StringComparison.Ordinal));
            Assert.Equal("Bold red italic\nSecond paragraph",
                binaryShape.TextBody.Text);
            Assert.False(binaryShape.TextBody.IsStyleTruncated,
                string.Join(Environment.NewLine, binary.Diagnostics));
            Assert.False(binaryShape.TextBody
                .HasUnprojectedCharacterFormatting);
            Assert.False(binaryShape.TextBody
                .HasUnprojectedParagraphFormatting);
            Assert.Equal(2, binaryShape.TextBody.ParagraphRuns.Count);
            LegacyPptParagraphRun firstParagraph = binaryShape.TextBody
                .ParagraphRuns[0];
            Assert.Equal((ushort)1, firstParagraph.IndentLevel);
            Assert.True(firstParagraph.HasBullet);
            Assert.Equal('•', firstParagraph.BulletCharacter);
            Assert.Equal("OfficeIMO Bullet",
                firstParagraph.BulletTypeface);
            Assert.Equal((short)120, firstParagraph.BulletSize);
            Assert.Equal("123456", firstParagraph.BulletColor);
            Assert.Equal(LegacyPptTextAlignment.Center,
                firstParagraph.Alignment);
            LegacyPptTextRuler ruler = Assert.IsType<LegacyPptTextRuler>(
                binaryShape.TextBody.Ruler);
            LegacyPptTextRulerLevel rulerLevel = Assert.Single(
                ruler.Levels, level => level.Level == 1);
            Assert.Equal((short)200, rulerLevel.LeftMargin);
            Assert.Equal((short)-100, rulerLevel.Indent);
            Assert.Equal((short)400, ruler.DefaultTabSize);
            LegacyPptTabStop tab = Assert.Single(ruler.TabStops);
            Assert.Equal((short)300, tab.Position);
            Assert.Equal(LegacyPptTabAlignment.Decimal, tab.Alignment);
            Assert.Equal((short)125, firstParagraph.LineSpacing);
            Assert.Equal((short)-100, firstParagraph.SpaceBefore);
            Assert.Equal(LegacyPptTextDirection.RightToLeft,
                firstParagraph.TextDirection);
            Assert.Collection(binaryShape.TextBody.CharacterRuns,
                run => AssertRun(run, "Bold red", 0, bold: true,
                    size: 32, color: "C00000"),
                run => AssertRun(run, " italic", 8, italic: true,
                    size: 24,
                    color: binary.Masters[0].ColorScheme.Accent1),
                run => AssertRun(run, "\n", 15, size: 20),
                run => AssertRun(run, "Second paragraph", 16,
                    underline: true, size: 20, color: "333333"));
            Assert.All(binaryShape.TextBody.CharacterRuns.Where(run =>
                    run.Text != "\n"),
                run => Assert.Equal("OfficeIMO Latin", run.Typeface));
            Assert.All(new[] { "OfficeIMO Bullet", "OfficeIMO Latin" },
                typeface => Assert.Contains(binary.Fonts,
                    font => font.Typeface == typeface));

            using var stream = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(stream);
            P.Shape projected = Assert.IsType<P.Shape>(Assert.Single(
                reopened.Slides[0].TextBoxes).Element);
            A.Paragraph[] paragraphs = projected.TextBody!
                .Elements<A.Paragraph>().ToArray();
            Assert.Equal(2, paragraphs.Length);
            Assert.Equal(1,
                paragraphs[0].ParagraphProperties!.Level!.Value);
            Assert.Equal("OfficeIMO Bullet", paragraphs[0]
                .ParagraphProperties!.GetFirstChild<A.BulletFont>()!
                .Typeface!.Value);
            Assert.True(paragraphs[0].Elements<A.Run>().First()
                .RunProperties!.Bold!.Value);
            Assert.Equal(A.TextUnderlineValues.Single,
                paragraphs[1].Elements<A.Run>().Single()
                    .RunProperties!.Underline!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_AuthorsTextFrameMarginsWrappingAnchorDirectionAndAutoFit() {
            byte[] bytes;
            using (PowerPointPresentation presentation = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints("Text frame", 30, 30, 240, 120)
                    .SetTextMarginsPoints(8, 4, 6, 2);
                textBox.TextVerticalAlignment =
                    A.TextAnchoringTypeValues.Center;
                textBox.TextDirection = A.TextVerticalValues.Vertical270;
                textBox.TextAutoFit = PowerPointTextAutoFit.Shape;
                A.BodyProperties body = Assert.IsType<P.Shape>(
                    textBox.Element).TextBody!.BodyProperties!;
                body.Wrap = A.TextWrappingValues.None;
                body.AnchorCenter = true;

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptShape binaryShape = Assert.Single(
                Assert.Single(binary.Slides).Shapes,
                shape => shape.Text == "Text frame");
            LegacyPptTextFrameProperties frame = binaryShape.TextFrame;
            Assert.Equal(101600, frame.LeftInsetEmus);
            Assert.Equal(50800, frame.TopInsetEmus);
            Assert.Equal(76200, frame.RightInsetEmus);
            Assert.Equal(25400, frame.BottomInsetEmus);
            Assert.Equal(2U, frame.WrapMode);
            Assert.Equal(4U, frame.AnchorMode);
            Assert.Equal(2U, frame.TextFlow);
            Assert.Equal(false, frame.AutoTextMargin);
            Assert.Equal(true, frame.FitShapeToText);

            using var stream = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(stream);
            PowerPointTextBox projected = Assert.Single(
                reopened.Slides[0].TextBoxes,
                textBox => textBox.Text == "Text frame");
            Assert.Equal(8D, projected.TextMarginLeftPoints);
            Assert.Equal(4D, projected.TextMarginTopPoints);
            Assert.Equal(6D, projected.TextMarginRightPoints);
            Assert.Equal(2D, projected.TextMarginBottomPoints);
            Assert.Equal(A.TextAnchoringTypeValues.Center,
                projected.TextVerticalAlignment);
            Assert.Equal(A.TextVerticalValues.Vertical270,
                projected.TextDirection);
            Assert.Equal(PowerPointTextAutoFit.Shape,
                projected.TextAutoFit);
            A.BodyProperties projectedBody = Assert.IsType<P.Shape>(
                projected.Element).TextBody!.BodyProperties!;
            Assert.Equal(A.TextWrappingValues.None,
                projectedBody.Wrap!.Value);
            Assert.True(projectedBody.AnchorCenter!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedTextFrameEdit_UsesIncrementalOfficeArtRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints("Editable frame", 30, 30, 240, 120)
                    .SetTextMarginsPoints(8, 4, 6, 2);
                textBox.TextVerticalAlignment =
                    A.TextAnchoringTypeValues.Center;
                textBox.TextDirection = A.TextVerticalValues.Vertical270;
                textBox.TextAutoFit = PowerPointTextAutoFit.Shape;
                A.BodyProperties body = Assert.IsType<P.Shape>(
                    textBox.Element).TextBody!.BodyProperties!;
                body.Wrap = A.TextWrappingValues.None;
                body.AnchorCenter = true;
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                PowerPointTextBox textBox = Assert.Single(
                    imported.Slides[0].TextBoxes,
                    candidate => candidate.Text == "Editable frame");
                LegacyPptShapeProjection projection = Assert.Single(
                    imported.LegacyPptProjectionMap!.Slides[0].Shapes,
                    candidate => candidate.OpenXmlShapeId == textBox.Id);
                Assert.True(projection.CanEditTextFrame);
                Assert.True(projection.TextFrameMatches(textBox));
                textBox.SetTextMarginsPoints(5, 7, 9, 11);
                textBox.TextVerticalAlignment =
                    A.TextAnchoringTypeValues.Bottom;
                textBox.TextDirection = A.TextVerticalValues.Vertical;
                textBox.TextAutoFit = PowerPointTextAutoFit.None;
                A.BodyProperties body = Assert.IsType<P.Shape>(
                    textBox.Element).TextBody!.BodyProperties!;
                body.Wrap = A.TextWrappingValues.Square;
                body.AnchorCenter = false;
                Assert.True(LegacyPptWriter.TryReadTextFrameForWrite(
                    textBox, out _, out string? frameReason), frameReason);
                Assert.False(projection.TextFrameMatches(textBox));

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptShape savedShape = Assert.Single(
                Assert.Single(saved.Slides).Shapes,
                shape => shape.Text == "Editable frame");
            LegacyPptTextFrameProperties frame = savedShape.TextFrame;
            Assert.Equal(63500, frame.LeftInsetEmus);
            Assert.Equal(88900, frame.TopInsetEmus);
            Assert.Equal(114300, frame.RightInsetEmus);
            Assert.Equal(139700, frame.BottomInsetEmus);
            Assert.Equal(0U, frame.WrapMode);
            Assert.Equal(2U, frame.AnchorMode);
            Assert.Equal(1U, frame.TextFlow);
            Assert.Equal(false, frame.AutoTextMargin);
            Assert.Equal(false, frame.FitShapeToText);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(reopenedInput);
            PowerPointTextBox projected = Assert.Single(
                reopened.Slides[0].TextBoxes,
                textBox => textBox.Text == "Editable frame");
            Assert.Equal(5D, projected.TextMarginLeftPoints);
            Assert.Equal(7D, projected.TextMarginTopPoints);
            Assert.Equal(9D, projected.TextMarginRightPoints);
            Assert.Equal(11D, projected.TextMarginBottomPoints);
            Assert.Equal(A.TextAnchoringTypeValues.Bottom,
                projected.TextVerticalAlignment);
            Assert.Equal(A.TextVerticalValues.Vertical,
                projected.TextDirection);
            Assert.Equal(PowerPointTextAutoFit.None,
                projected.TextAutoFit);
            A.BodyProperties projectedBody = Assert.IsType<P.Shape>(
                projected.Element).TextBody!.BodyProperties!;
            Assert.Equal(A.TextWrappingValues.Square,
                projectedBody.Wrap!.Value);
            Assert.False(projectedBody.AnchorCenter?.Value ?? false);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedPresentation_AppendedSlideAuthorsRichTextAndFonts() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBox("Existing slide");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                PowerPointTextBox textBox = imported.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 40, 40, 320, 120);
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                shape.TextBody = new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.ParagraphProperties(
                            new A.CharacterBullet { Char = "•" }) {
                            Alignment = A.TextAlignmentTypeValues.Center
                        },
                        new A.Run(
                            new A.RunProperties(
                                new A.LatinFont {
                                    Typeface = "OfficeIMO Appended"
                                }) {
                                Bold = true,
                                FontSize = 2600
                            },
                            new A.Text("Appended rich text"))));

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptShape savedShape = Assert.Single(saved.Slides[1].Shapes,
                shape => shape.Text == "Appended rich text");
            LegacyPptCharacterRun savedRun = Assert.Single(
                savedShape.TextBody.CharacterRuns,
                run => run.Text == "Appended rich text");
            Assert.True(savedRun.Bold);
            Assert.Equal((short)26, savedRun.FontSizePoints);
            Assert.Equal("OfficeIMO Appended", savedRun.Typeface);
            LegacyPptParagraphRun paragraph = Assert.Single(
                savedShape.TextBody.ParagraphRuns);
            Assert.True(paragraph.HasBullet);
            Assert.Equal('•', paragraph.BulletCharacter);
            Assert.Equal(LegacyPptTextAlignment.Center,
                paragraph.Alignment);
            Assert.Contains(saved.Fonts,
                font => font.Typeface == "OfficeIMO Appended");
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(reopenedInput);
            P.Shape reopenedShape = Assert.IsType<P.Shape>(Assert.Single(
                reopened.Slides[1].TextBoxes,
                candidate => candidate.Text == "Appended rich text").Element);
            A.Run reopenedRun = Assert.Single(reopenedShape.TextBody!
                .Descendants<A.Run>());
            Assert.True(reopenedRun.RunProperties!.Bold!.Value);
            Assert.Equal("OfficeIMO Appended", reopenedRun.RunProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedRichTextGeometryEdit_PreservesCharacterRuns() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(RichTextFixturePath);
            LegacyPptShape originalShape = Assert.Single(Assert.Single(original.Slides).Shapes,
                shape => shape.Text.Contains("Bold red", StringComparison.Ordinal));
            using PowerPointPresentation presentation = PowerPointPresentation.Load(RichTextFixturePath);
            PowerPointTextBox textBox = Assert.Single(presentation.Slides[0].TextBoxes);

            textBox.Left += 15875;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedShape = Assert.Single(Assert.Single(saved.Slides).Shapes,
                shape => shape.Text.Contains("Bold red", StringComparison.Ordinal));
            Assert.Equal(originalShape.Bounds.Left + 10, savedShape.Bounds.Left);
            Assert.Equal(originalShape.TextBody.CharacterRuns.Select(DescribeRun),
                savedShape.TextBody.CharacterRuns.Select(DescribeRun));
            Assert.Equal(original.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedRichTextFormattingAndLengthEdit_UsesIncrementalStyleRewrite() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                RichTextFixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(RichTextFixturePath);
            P.Shape shape = Assert.IsType<P.Shape>(Assert.Single(
                presentation.Slides[0].TextBoxes).Element);
            A.Run first = shape.TextBody!.Descendants<A.Run>().First();
            first.RunProperties!.Bold = false;
            first.RunProperties.FontSize = 3000;
            first.RunProperties.RemoveAllChildren<A.LatinFont>();
            first.RunProperties.Append(new A.LatinFont {
                Typeface = "OfficeIMO Edited"
            });
            first.Text!.Text += "!";
            shape.TextBody.Elements<A.Paragraph>().First()
                .ParagraphProperties!.Alignment =
                A.TextAlignmentTypeValues.Right;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();

            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedShape = Assert.Single(
                Assert.Single(saved.Slides).Shapes,
                candidate => candidate.Text.StartsWith("Bold red!",
                    StringComparison.Ordinal));
            LegacyPptCharacterRun savedFirst = savedShape.TextBody
                .CharacterRuns[0];
            Assert.False(savedFirst.Bold);
            Assert.Equal((short)30, savedFirst.FontSizePoints);
            Assert.Equal("OfficeIMO Edited", savedFirst.Typeface);
            Assert.Equal(LegacyPptTextAlignment.Right,
                savedShape.TextBody.ParagraphRuns[0].Alignment);
            Assert.Contains(saved.Fonts,
                font => font.Typeface == "OfficeIMO Edited");
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedMasterTextStyleEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(RichTextFixturePath);
            P.SlideMaster slideMaster = presentation.Slides[0].SlidePart.SlideLayoutPart!
                .SlideMasterPart!.SlideMaster!;
            A.Level1ParagraphProperties title = Assert.IsType<A.Level1ParagraphProperties>(
                slideMaster.TextStyles!.TitleStyle!.FirstChild);
            title.GetFirstChild<A.DefaultRunProperties>()!.FontSize = 4500;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void ImportedTextRulerEdit_UsesIncrementalRulerRewrite() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                FixturePath);
            LegacyPptShape originalTitle = Assert.Single(
                Assert.Single(original.Slides).Shapes,
                candidate => candidate.Text
                    == "OfficeIMO PowerPoint Basics");
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            P.Shape title = Assert.IsType<P.Shape>(presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "OfficeIMO PowerPoint Basics").Element);
            A.TabStop tab = title.TextBody!.Descendants<A.TabStop>().First();
            tab.Position = tab.Position!.Value + 15875;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();

            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedTitle = Assert.Single(
                Assert.Single(saved.Slides).Shapes,
                candidate => candidate.Text
                    == "OfficeIMO PowerPoint Basics");
            Assert.Equal(originalTitle.TextBody.Ruler!.TabStops[0].Position
                    + 10,
                savedTitle.TextBody.Ruler!.TabStops[0].Position);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
        }

        [Fact]
        public void TextStyleReader_RejectsTruncatedRunWithoutEscapingBounds() {
            byte[] payload = { 0x02, 0x00, 0x00, 0x00 };
            var record = new LegacyPptRecord(payload, 0, 0, 0, 0x0FA1, 0, payload.Length);

            LegacyPptTextBody result = LegacyPptTextStyleReader.Read("A", 1, record,
                colorScheme: null);

            Assert.True(result.HasStyleRecord);
            Assert.True(result.IsStyleTruncated);
            Assert.Empty(result.CharacterRuns);
        }

        [Fact]
        public void TextStyleReader_ProjectsExplicitOverridesSchemeColorAndBaseline() {
            byte[] payload = {
                0x03, 0x00, 0x00, 0x00, // TextPFRun count, including terminal character
                0x00, 0x00,             // indent level
                0x00, 0x00, 0x00, 0x00, // paragraph masks
                0x01, 0x00, 0x00, 0x00, // first TextCFRun count
                0x07, 0x00, 0x0E, 0x00, // bold, italic, underline, size, color, position masks
                0x02, 0x00,             // italic true; bold and underline explicitly false
                0x12, 0x00,             // 18 points
                0x00, 0x00, 0x00, 0x05, // scheme color index 5
                0xE7, 0xFF,             // -25 percent baseline position
                0x02, 0x00, 0x00, 0x00, // second TextCFRun covers B and terminal character
                0x01, 0x00, 0x00, 0x00, // bold mask
                0x01, 0x00              // bold true
            };
            var record = new LegacyPptRecord(payload, 0, 0, 0, 0x0FA1, 0, payload.Length);
            var scheme = new LegacyPptColorScheme(new[] {
                "FFFFFF", "000000", "777777", "111111",
                "EEEEEE", "ABCDEF", "123456", "654321"
            });

            LegacyPptTextBody result = LegacyPptTextStyleReader.Read("AB", 2, record, scheme);

            Assert.False(result.IsStyleTruncated);
            Assert.False(result.HasParagraphFormatting);
            Assert.False(result.HasUnprojectedCharacterFormatting);
            Assert.Collection(result.CharacterRuns,
                first => {
                    AssertRun(first, "A", 0, bold: false, italic: true, underline: false,
                        size: 18, color: "ABCDEF");
                    Assert.Equal((byte)5, first.ColorSchemeIndex);
                    Assert.Equal((short)-25, first.BaselinePositionPercent);
                },
                second => AssertRun(second, "B", 1, bold: true));

            P.TextBody projected = LegacyPptTextProjection.CreateTextBody(result);
            A.Run[] runs = Assert.Single(projected.Elements<A.Paragraph>()).Elements<A.Run>().ToArray();
            Assert.False(runs[0].RunProperties!.Bold!.Value);
            Assert.True(runs[0].RunProperties.Italic!.Value);
            Assert.Equal(A.TextUnderlineValues.None, runs[0].RunProperties.Underline!.Value);
            Assert.Equal(1800, runs[0].RunProperties.FontSize!.Value);
            Assert.Equal(-25000, runs[0].RunProperties.Baseline!.Value);
            Assert.Equal(A.SchemeColorValues.Accent1, runs[0].RunProperties
                .GetFirstChild<A.SolidFill>()!.SchemeColor!.Val!.Value);
            Assert.True(runs[1].RunProperties!.Bold!.Value);
        }

        [Fact]
        public void TextStyleReader_ProjectsParagraphSpacingBulletsDirectionAndWrapping() {
            byte[] payload;
            using (var stream = new MemoryStream()) {
                using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
                    writer.Write(2U); // TextPFRun covers A and the terminal character
                    writer.Write((ushort)2);
                    uint masks = 0x000000FFU | (1U << 11) | (1U << 12) | (1U << 13)
                        | (1U << 14) | (1U << 16) | (7U << 17) | (1U << 21);
                    writer.Write(masks);
                    writer.Write((ushort)0x000F); // bullet enabled with font, color, and size
                    writer.Write((ushort)'•');
                    writer.Write((ushort)1);      // Arial
                    writer.Write((short)80);      // 80 percent
                    writer.Write((byte)0xAA);
                    writer.Write((byte)0xBB);
                    writer.Write((byte)0xCC);
                    writer.Write((byte)0xFE);
                    writer.Write((ushort)2);      // right aligned
                    writer.Write((short)150);     // 150 percent line spacing
                    writer.Write((short)-80);     // 10 points before
                    writer.Write((short)25);      // 25 percent after
                    writer.Write((ushort)2);      // centered within line height
                    writer.Write((ushort)0x0003); // East Asian break + word-only wrapping
                    writer.Write((ushort)1);      // right to left
                    writer.Write(2U);             // TextCFRun covers A and terminal character
                    writer.Write(0U);
                }
                payload = stream.ToArray();
            }
            var record = new LegacyPptRecord(payload, 0, 0, 0, 0x0FA1, 0, payload.Length);
            var arial = new LegacyPptFont(1, "Arial", 0, false, false, false,
                true, false, 0x22, false);

            LegacyPptTextBody result = LegacyPptTextStyleReader.Read("A", 1, record,
                colorScheme: null, fonts: new Dictionary<ushort, LegacyPptFont> { [1] = arial });

            Assert.False(result.IsStyleTruncated);
            Assert.False(result.HasUnprojectedParagraphFormatting);
            LegacyPptParagraphRun paragraph = Assert.Single(result.ParagraphRuns);
            Assert.Equal((ushort)2, paragraph.IndentLevel);
            Assert.True(paragraph.HasBullet);
            Assert.Equal('•', paragraph.BulletCharacter);
            Assert.Equal("Arial", paragraph.BulletTypeface);
            Assert.Equal((short)80, paragraph.BulletSize);
            Assert.Equal("AABBCC", paragraph.BulletColor);
            Assert.Equal(LegacyPptTextAlignment.Right, paragraph.Alignment);
            Assert.Equal((short)150, paragraph.LineSpacing);
            Assert.Equal((short)-80, paragraph.SpaceBefore);
            Assert.Equal((short)25, paragraph.SpaceAfter);
            Assert.Equal(LegacyPptFontAlignment.Center, paragraph.FontAlignment);
            Assert.True(paragraph.CharacterWrap);
            Assert.True(paragraph.WordWrap);
            Assert.False(paragraph.Overflow);
            Assert.Equal(LegacyPptTextDirection.RightToLeft, paragraph.TextDirection);

            A.ParagraphProperties properties = Assert.Single(
                LegacyPptTextProjection.CreateTextBody(result).Elements<A.Paragraph>())
                .ParagraphProperties!;
            Assert.Equal(2, properties.Level!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Right, properties.Alignment!.Value);
            Assert.Equal(A.TextFontAlignmentValues.Center, properties.FontAlignment!.Value);
            Assert.True(properties.RightToLeft!.Value);
            Assert.True(properties.EastAsianLineBreak!.Value);
            Assert.False(properties.LatinLineBreak!.Value);
            Assert.False(properties.Height!.Value);
            Assert.Equal(150000, properties.LineSpacing!.SpacingPercent!.Val!.Value);
            Assert.Equal(1000, properties.SpaceBefore!.SpacingPoints!.Val!.Value);
            Assert.Equal(25000, properties.SpaceAfter!.SpacingPercent!.Val!.Value);
            Assert.Equal("AABBCC", properties.GetFirstChild<A.BulletColor>()!
                .RgbColorModelHex!.Val!.Value);
            Assert.Equal(80000, properties.GetFirstChild<A.BulletSizePercentage>()!.Val!.Value);
            Assert.Equal("Arial", properties.GetFirstChild<A.BulletFont>()!.Typeface!.Value);
            Assert.Equal("•", properties.GetFirstChild<A.CharacterBullet>()!.Char!.Value);
        }

        [Fact]
        public void TextRulerReader_ProjectsTabsMarginsAndIndentation() {
            byte[] payload;
            using (var stream = new MemoryStream()) {
                using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
                    uint mask = (1U << 0) | (1U << 1) | (1U << 2)
                        | (1U << 3) | (1U << 4) | (1U << 8) | (1U << 9);
                    writer.Write(mask);
                    writer.Write((short)2);       // ruler level count
                    writer.Write((short)720);     // default tab size
                    writer.Write((ushort)2);      // tab-stop count
                    writer.Write((short)576);
                    writer.Write((ushort)LegacyPptTabAlignment.Left);
                    writer.Write((short)1152);
                    writer.Write((ushort)LegacyPptTabAlignment.Decimal);
                    writer.Write((short)144);     // level 0 left margin
                    writer.Write((short)-72);     // level 0 first-line indent
                    writer.Write((short)432);     // level 1 left margin
                    writer.Write((short)288);     // level 1 first-line indent
                }
                payload = stream.ToArray();
            }
            var record = new LegacyPptRecord(payload, 0, 0, 0, 0x0FA6, 0, payload.Length);

            LegacyPptTextRuler ruler = Assert.IsType<LegacyPptTextRuler>(
                LegacyPptTextRulerReader.Read(record, out bool truncated));

            Assert.False(truncated);
            Assert.Equal((short)2, ruler.LevelCount);
            Assert.Equal((short)720, ruler.DefaultTabSize);
            Assert.Collection(ruler.Levels,
                level => {
                    Assert.Equal((ushort)0, level.Level);
                    Assert.Equal((short)144, level.LeftMargin);
                    Assert.Equal((short)-72, level.Indent);
                },
                level => {
                    Assert.Equal((ushort)1, level.Level);
                    Assert.Equal((short)432, level.LeftMargin);
                    Assert.Equal((short)288, level.Indent);
                });
            Assert.Collection(ruler.TabStops,
                tab => {
                    Assert.Equal((short)576, tab.Position);
                    Assert.Equal(LegacyPptTabAlignment.Left, tab.Alignment);
                },
                tab => {
                    Assert.Equal((short)1152, tab.Position);
                    Assert.Equal(LegacyPptTabAlignment.Decimal, tab.Alignment);
                });

            LegacyPptTextBody body = LegacyPptTextStyleReader.Read("A", 1, styleRecord: null,
                colorScheme: null, textType: LegacyPptTextType.Other, ruler: ruler,
                hasRulerRecord: true);
            A.ParagraphProperties properties = Assert.Single(
                LegacyPptTextProjection.CreateTextBody(body).Elements<A.Paragraph>())
                .ParagraphProperties!;
            Assert.Equal(228600, properties.LeftMargin!.Value);
            Assert.Equal(-114300, properties.Indent!.Value);
            Assert.Equal(1143000, properties.DefaultTabSize!.Value);
            Assert.Collection(properties.GetFirstChild<A.TabStopList>()!.Elements<A.TabStop>(),
                tab => {
                    Assert.Equal(914400, tab.Position!.Value);
                    Assert.Equal(A.TextTabAlignmentValues.Left, tab.Alignment!.Value);
                },
                tab => {
                    Assert.Equal(1828800, tab.Position!.Value);
                    Assert.Equal(A.TextTabAlignmentValues.Decimal, tab.Alignment!.Value);
                });
        }

        [Fact]
        public void TextRulerReader_RejectsReservedMaskBitsWithoutEscapingBounds() {
            byte[] payload = { 0x00, 0x00, 0x00, 0x80 };
            var record = new LegacyPptRecord(payload, 0, 0, 0, 0x0FA6, 0, payload.Length);

            LegacyPptTextRuler? ruler = LegacyPptTextRulerReader.Read(record, out bool truncated);

            Assert.Null(ruler);
            Assert.True(truncated);
        }

        [Fact]
        public void TextMasterStyleReader_DecodesExplicitCenterStyleLevels() {
            byte[] payload;
            using (var stream = new MemoryStream()) {
                using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
                    writer.Write((ushort)2);       // style-level count
                    writer.Write((ushort)1);       // explicit level for text types 5 through 8
                    writer.Write(0U);              // empty TextPFException
                    writer.Write(1U << 17);        // character size
                    writer.Write((short)18);
                    writer.Write((ushort)0);       // explicit level zero
                    writer.Write(1U << 11);        // paragraph alignment
                    writer.Write((ushort)1);       // centered
                    writer.Write(0U);              // empty TextCFException
                }
                payload = stream.ToArray();
            }
            var record = new LegacyPptRecord(payload, 0, 0, 5, 0x0FA3, 0, payload.Length);

            LegacyPptTextMasterStyle style = Assert.IsType<LegacyPptTextMasterStyle>(
                LegacyPptTextMasterStyleReader.Read(record, colorScheme: null, fonts: null));

            Assert.False(style.IsTruncated);
            Assert.False(style.HasUnprojectedFormatting);
            Assert.Equal(LegacyPptTextType.CenterBody, style.TextType);
            Assert.Collection(style.Levels,
                level => {
                    Assert.Equal((ushort)0, level.Level);
                    Assert.Equal(LegacyPptTextAlignment.Center,
                        level.ParagraphProperties.Alignment);
                },
                level => {
                    Assert.Equal((ushort)1, level.Level);
                    Assert.Equal((short)18, level.CharacterProperties.FontSizePoints);
                });
        }

        [Fact]
        public void TextMasterStyleReader_RejectsInvalidExplicitLevelWithoutEscapingBounds() {
            byte[] payload = {
                0x01, 0x00, // one level
                0x01, 0x00, // invalid: level must be less than cLevels
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00
            };
            var record = new LegacyPptRecord(payload, 0, 0, 5, 0x0FA3, 0, payload.Length);

            LegacyPptTextMasterStyle style = Assert.IsType<LegacyPptTextMasterStyle>(
                LegacyPptTextMasterStyleReader.Read(record, colorScheme: null, fonts: null));

            Assert.True(style.IsTruncated);
            Assert.Empty(style.Levels);
        }

        [Fact]
        public void NativeWriter_AuthorsPpt9AutomaticNumberingAcrossRunGroups() {
            byte[] bytes;
            using (PowerPointPresentation presentation = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 180);
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                shape.TextBody = new P.TextBody(new A.BodyProperties(),
                    new A.ListStyle(),
                    CreateNumberedParagraph("Fourth", 0,
                        A.TextAutoNumberSchemeValues.ArabicPeriod, 4),
                    new A.Paragraph(new A.Run(new A.Text("Plain"))),
                    CreateNumberedParagraph("Third letter", 2,
                        A.TextAutoNumberSchemeValues
                            .AlphaLowerCharacterParenR, 3),
                    CreateNumberedParagraph("Roman", 0,
                        A.TextAutoNumberSchemeValues
                            .RomanUpperCharacterPeriod, 9));

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptShape binaryShape = Assert.Single(
                Assert.Single(binary.Slides).Shapes,
                shape => shape.Text.StartsWith("Fourth",
                    StringComparison.Ordinal));
            Assert.True(binaryShape.TextBody.HasStyle9Record);
            Assert.False(binaryShape.TextBody.IsStyle9Truncated,
                string.Join(Environment.NewLine, binary.Diagnostics));
            Assert.False(binaryShape.TextBody
                .HasUnprojectedParagraphFormatting);
            Assert.Collection(binaryShape.TextBody.ParagraphRuns,
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.ArabicPeriod, 4),
                paragraph => {
                    Assert.Null(paragraph.HasAutoNumber);
                    Assert.Null(paragraph.AutoNumberScheme);
                },
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.AlphaLowerParenRight, 3),
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.RomanUpperPeriod, 9));
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 0);
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 1);
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 2);
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 3);

            using var stream = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(stream);
            A.Paragraph[] paragraphs = Assert.IsType<P.Shape>(Assert.Single(
                    reopened.Slides[0].TextBoxes).Element).TextBody!
                .Elements<A.Paragraph>().ToArray();
            Assert.Equal(4, paragraphs.Length);
            Assert.Equal(A.TextAutoNumberSchemeValues.ArabicPeriod,
                paragraphs[0].ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!.Type!.Value);
            Assert.Null(paragraphs[1].ParagraphProperties);
            Assert.Equal(3, paragraphs[2].ParagraphProperties!
                .GetFirstChild<A.AutoNumberedBullet>()!.StartAt!.Value);
            Assert.Equal(A.TextAutoNumberSchemeValues
                    .RomanUpperCharacterPeriod,
                paragraphs[3].ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!.Type!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedPlainText_AddsPpt9NumberingWithAppendOnlyRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBox("Number me");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                P.Shape shape = Assert.IsType<P.Shape>(Assert.Single(
                    imported.Slides[0].TextBoxes).Element);
                A.Paragraph paragraph = Assert.Single(shape.TextBody!
                    .Elements<A.Paragraph>());
                paragraph.ParagraphProperties = new A.ParagraphProperties(
                    new A.AutoNumberedBullet {
                        Type = A.TextAutoNumberSchemeValues.ArabicParenBoth,
                        StartAt = 7
                    });
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptParagraphRun paragraphRun = Assert.Single(
                Assert.Single(Assert.Single(saved.Slides).Shapes,
                    shape => shape.Text == "Number me").TextBody
                    .ParagraphRuns);
            AssertAutoNumber(paragraphRun,
                LegacyPptAutoNumberScheme.ArabicParenBoth, 7);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedAutomaticNumbering_CanChangeAndRemoveEntries() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 120);
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        CreateNumberedParagraph("Change", 0,
                            A.TextAutoNumberSchemeValues.ArabicPeriod, 2),
                        CreateNumberedParagraph("Remove", 0,
                            A.TextAutoNumberSchemeValues.ArabicPeriod, 3));
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                A.Paragraph[] paragraphs = Assert.IsType<P.Shape>(
                        Assert.Single(imported.Slides[0].TextBoxes).Element)
                    .TextBody!.Elements<A.Paragraph>().ToArray();
                A.AutoNumberedBullet first = paragraphs[0]
                    .ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!;
                first.Type = A.TextAutoNumberSchemeValues
                    .AlphaUpperCharacterParenBoth;
                first.StartAt = 5;
                paragraphs[1].ParagraphProperties!
                    .RemoveAllChildren<A.AutoNumberedBullet>();
                paragraphs[1].ParagraphProperties!
                    .Append(new A.NoBullet());

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptShape shape = Assert.Single(
                Assert.Single(saved.Slides).Shapes,
                candidate => candidate.Text.StartsWith("Change",
                    StringComparison.Ordinal));
            AssertAutoNumber(shape.TextBody.ParagraphRuns[0],
                LegacyPptAutoNumberScheme.AlphaUpperParenBoth, 5);
            Assert.False(shape.TextBody.ParagraphRuns[1].HasBullet);
            Assert.Null(shape.TextBody.ParagraphRuns[1].HasAutoNumber);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedAutomaticNumbering_RemovesLastPpt9Tag() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 120);
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        CreateNumberedParagraph("Remove all", 0,
                            A.TextAutoNumberSchemeValues.ArabicPeriod, 1));
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            Assert.True(Assert.Single(Assert.Single(original.Slides).Shapes)
                .TextBody.HasStyle9Record);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                A.ParagraphProperties properties = Assert.IsType<P.Shape>(
                        Assert.Single(imported.Slides[0].TextBoxes).Element)
                    .TextBody!.Elements<A.Paragraph>().Single()
                    .ParagraphProperties!;
                properties.RemoveAllChildren<A.AutoNumberedBullet>();
                properties.Append(new A.NoBullet());
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptShape shape = Assert.Single(
                Assert.Single(saved.Slides).Shapes);
            Assert.False(shape.TextBody.HasStyle9Record);
            Assert.False(Assert.Single(shape.TextBody.ParagraphRuns)
                .HasBullet);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void NormalAutoFit_IsExplicitlyBlockedAsNonRepresentable() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointTextBox textBox = presentation.AddSlide(
                    P.SlideLayoutValues.Blank)
                .AddTextBox("Shrink me");
            textBox.TextAutoFit = PowerPointTextAutoFit.Normal;

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            LegacyPptWriteFinding finding = Assert.Single(
                preflight.Findings,
                item => item.Code == "PPT-WRITE-RICH-TEXT");
            Assert.Contains("no lossless classic binary PowerPoint mapping",
                finding.Description, StringComparison.Ordinal);
        }

        private static A.Paragraph CreateNumberedParagraph(string text,
            int level, A.TextAutoNumberSchemeValues scheme, int startAt) =>
            new(new A.ParagraphProperties(
                    new A.AutoNumberedBullet {
                        Type = scheme,
                        StartAt = startAt
                    }) {
                    Level = level
                },
                new A.Run(new A.Text(text)));

        private static void AssertAutoNumber(
            LegacyPptParagraphRun paragraph,
            LegacyPptAutoNumberScheme scheme, short startAt) {
            Assert.True(paragraph.HasAutoNumber);
            Assert.Equal(scheme, paragraph.AutoNumberScheme);
            Assert.Equal(startAt, paragraph.AutoNumberStartAt);
        }

        private static void AssertRun(LegacyPptCharacterRun run, string text, int start,
            bool? bold = null, bool? italic = null, bool? underline = null,
            short? size = null, string? color = null) {
            Assert.Equal(text, run.Text);
            Assert.Equal(start, run.Start);
            Assert.Equal(text.Length, run.Length);
            Assert.Equal(bold, run.Bold);
            Assert.Equal(italic, run.Italic);
            Assert.Equal(underline, run.Underline);
            Assert.Equal(size, run.FontSizePoints);
            Assert.Equal(color, run.Color);
        }

        private static string? GetRunColor(A.Run run) => run.RunProperties?
            .GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value;

        private static string DescribeRun(LegacyPptCharacterRun run) => string.Join("|",
            run.Start, run.Length, run.Text, run.Bold, run.Italic, run.Underline,
            run.FontIndex, run.Typeface, run.FontSizePoints, run.Color, run.ColorSchemeIndex,
            run.BaselinePositionPercent);
    }
}
