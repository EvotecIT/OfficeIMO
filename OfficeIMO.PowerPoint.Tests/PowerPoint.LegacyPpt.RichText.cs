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
            Assert.True(textBody.HasUnprojectedCharacterFormatting);
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
                Assert.True(run.HasUnprojectedFormatting);
            });
            Assert.Contains(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-PARAGRAPH-PARTIAL");
            Assert.Contains(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-CHARACTER-PARTIAL");
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

            Assert.Equal(2, paragraphs.Length);
            A.Run[] firstParagraph = paragraphs[0].Elements<A.Run>().ToArray();
            Assert.Equal(5, firstParagraph.Length);
            Assert.Equal("Bold red", firstParagraph[0].Text!.Text);
            Assert.True(firstParagraph[0].RunProperties!.Bold!.Value);
            Assert.Equal(3200, firstParagraph[0].RunProperties.FontSize!.Value);
            Assert.Equal("C00000", GetRunColor(firstParagraph[0]));
            Assert.True(firstParagraph[2].RunProperties!.Italic!.Value);
            Assert.Equal(2600, firstParagraph[2].RunProperties.FontSize!.Value);
            Assert.Equal("008000", GetRunColor(firstParagraph[2]));
            Assert.Equal(A.TextUnderlineValues.Single,
                firstParagraph[4].RunProperties!.Underline!.Value);
            Assert.Equal(2200, firstParagraph[4].RunProperties.FontSize!.Value);
            Assert.Equal("0000C0", GetRunColor(firstParagraph[4]));
            A.Run secondParagraph = Assert.Single(paragraphs[1].Elements<A.Run>());
            Assert.Equal("Regular second paragraph", secondParagraph.Text!.Text);
            Assert.Equal(2000, secondParagraph.RunProperties!.FontSize!.Value);
            Assert.Equal("333333", GetRunColor(secondParagraph));
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
            Assert.Equal("ABCDEF", GetRunColor(runs[0]));
            Assert.True(runs[1].RunProperties!.Bold!.Value);
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
            run.FontIndex, run.FontSizePoints, run.Color, run.ColorSchemeIndex,
            run.BaselinePositionPercent);
    }
}
