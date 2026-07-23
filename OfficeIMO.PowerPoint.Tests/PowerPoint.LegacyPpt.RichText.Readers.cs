using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
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
        public void TextProjection_DropsInvalidXmlCharactersFromLegacyFontsAndBullets() {
            var characterRun = new LegacyPptCharacterRun(
                0, 1, "A", bold: null, italic: null, underline: null,
                shadow: null, farEastHint: null, kumi: null, emboss: null,
                fontIndex: 1, oldEastAsianFontIndex: 2, ansiFontIndex: 3,
                symbolFontIndex: 4, typeface: "Primary\u0001Font",
                oldEastAsianTypeface: "East\uD800Font",
                ansiTypeface: "Ansi\u0002Font",
                symbolTypeface: "Symbol\uDC00Font", fontSizePoints: null,
                color: null, colorSchemeIndex: null,
                baselinePositionPercent: null,
                hasUnprojectedFormatting: false);
            var paragraphRun = new LegacyPptParagraphRun(
                0, 1, 0, hasBullet: true, bulletHasFont: true,
                bulletHasColor: null, bulletHasSize: null,
                bulletCharacter: '\uD800', bulletFontIndex: 5,
                bulletTypeface: "Bullet\u0003Font", bulletSize: null,
                bulletColor: null, bulletColorSchemeIndex: null,
                alignment: null, lineSpacing: null, spaceBefore: null,
                spaceAfter: null, fontAlignment: null,
                characterWrap: null, wordWrap: null, overflow: null,
                textDirection: null, hasUnprojectedFormatting: false);
            var body = new LegacyPptTextBody(
                "A", new[] { characterRun }, new[] { paragraphRun },
                hasStyleRecord: true,
                hasUnprojectedCharacterFormatting: false,
                hasUnprojectedParagraphFormatting: false);

            A.Paragraph projected = Assert.Single(
                LegacyPptTextProjection.CreateTextBody(body)
                    .Elements<A.Paragraph>());
            A.ParagraphProperties paragraphProperties =
                projected.ParagraphProperties!;
            Assert.Equal("BulletFont", paragraphProperties
                .GetFirstChild<A.BulletFont>()!.Typeface!.Value);
            Assert.Null(paragraphProperties.GetFirstChild<A.CharacterBullet>());
            A.RunProperties runProperties = Assert.Single(
                projected.Elements<A.Run>()).RunProperties!;
            Assert.Equal("AnsiFont", runProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal("EastFont", runProperties
                .GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
            Assert.Equal("SymbolFont", runProperties
                .GetFirstChild<A.SymbolFont>()!.Typeface!.Value);
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
        public void TextPropertyReader_RejectsOversizedTabStopCountBeforeAllocation() {
            byte[] payload = { 0x01, 0x10 }; // 4,097 entries
            var record = new LegacyPptRecord(payload, 0, 0, 0,
                0x0FA6, 0, payload.Length);
            var cursor = new LegacyPptTextPropertyCursor(record,
                "oversized tab stops");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => LegacyPptTextPropertyReader.ReadTabStops(cursor));

            Assert.Contains("4096", exception.Message,
                StringComparison.Ordinal);
            Assert.Equal(2, cursor.Offset);
        }

        [Fact]
        public void TextStyle9Reader_StopsAtConfiguredEntryLimit() {
            byte[] payload = new byte[24]; // two empty 12-byte StyleTextProp9 entries
            var record = new LegacyPptRecord(payload, 0, 0, 0,
                0x0FAC, 0, payload.Length);

            LegacyPptTextBody result = LegacyPptTextStyle9Reader.Apply(
                LegacyPptTextBody.Plain("A"), record,
                maximumEntryCount: 1);

            Assert.True(result.HasStyle9Record);
            Assert.True(result.IsStyle9Truncated);
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

    }
}
