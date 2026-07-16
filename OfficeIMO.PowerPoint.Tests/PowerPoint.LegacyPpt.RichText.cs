using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
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
        public void ImportedRichTextFormattingEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(RichTextFixturePath);
            P.Shape shape = Assert.IsType<P.Shape>(Assert.Single(
                presentation.Slides[0].TextBoxes).Element);
            shape.TextBody!.Descendants<A.Run>().First().RunProperties!.Bold = false;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
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
        public void ImportedTextRulerEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            P.Shape title = Assert.IsType<P.Shape>(presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "OfficeIMO PowerPoint Basics").Element);
            A.TabStop tab = title.TextBody!.Descendants<A.TabStop>().First();
            tab.Position = tab.Position!.Value + 1588;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
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
