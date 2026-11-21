using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_EmbeddedFontPresent() {
            string fontPath = Path.Combine(_directoryWithFiles, "DummyFont.ttf");
            File.WriteAllText(fontPath, "dummy");
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithEmbeddedFont.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.EmbedFont(fontPath);
                document.Save();
            }

            using var word = WordprocessingDocument.Open(filePath, false);
            var fontTablePart = word.MainDocumentPart?.FontTablePart;
            Assert.NotNull(fontTablePart);
            Assert.True(fontTablePart!.FontParts.Any());
            File.Delete(fontPath);
        }

        [Fact]
        public void Test_EmbedFontWithStyleRegistersStyle() {
            string fontPath = Path.Combine(_directoryWithFiles, "DummyFont.ttf");
            File.WriteAllText(fontPath, "dummy");
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithFontStyle.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.EmbedFont(fontPath, "DejaVuStyle", "DejaVu Style");
                document.AddParagraph("Test").SetStyleId("DejaVuStyle");
                document.Save();
            }

            using var word = WordprocessingDocument.Open(filePath, false);
            var stylePart = word.MainDocumentPart?.StyleDefinitionsPart;
            Assert.NotNull(stylePart);
            var styles = stylePart!.Styles;
            Assert.NotNull(styles);
            Assert.NotNull(styles!.Elements<Style>().FirstOrDefault(s => s.StyleId == "DejaVuStyle"));
            File.Delete(fontPath);
        }

        [Fact]
        public void Test_EmbedFontWithTakenStyleNameLeavesFontTableUnchanged() {
            string fontPath = Path.Combine(_directoryWithFiles, "DummyFontTakenName.ttf");
            File.WriteAllText(fontPath, "dummy");
            using (var document = WordDocument.Create()) {
                document.Styles.AddParagraphStyle("Existing", "Brand Font");
                int fontsBefore = document._wordprocessingDocument.MainDocumentPart!.FontTablePart?.Fonts?.Elements<Font>().Count() ?? 0;

                Assert.Throws<ArgumentException>(() => document.EmbedFont(fontPath, "BrandFont", "Brand Font"));
                document.Styles.AddCharacterStyle("BrandCharacter");
                Assert.Throws<ArgumentException>(() => document.EmbedFont(fontPath, "BrandCharacter"));

                Assert.False(document.Styles.Contains("BrandFont"));
                Assert.Equal(fontsBefore, document._wordprocessingDocument.MainDocumentPart!.FontTablePart?.Fonts?.Elements<Font>().Count() ?? 0);
            }
            File.Delete(fontPath);
        }

        [Theory]
        [InlineData(0x0000, true)]
        [InlineData(0x0004, true)]
        [InlineData(0x0006, true)]
        [InlineData(0x0002, false)]
        [InlineData(0x0200, false)]
        public void Test_EmbedFontHonoursTheFontEmbeddingPermission(int fsType, bool embedded) {
            byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "NotoSansSC-BaselineSubset.ttf"));
            int tableCount = (font[4] << 8) | font[5];
            int os2 = Enumerable.Range(0, tableCount).Select(index => 12 + index * 16)
                .Where(record => System.Text.Encoding.ASCII.GetString(font, record, 4) == "OS/2")
                .Select(record => (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11]).Single();
            font[os2 + 8] = (byte)(fsType >> 8);
            font[os2 + 9] = (byte)fsType;
            string fontPath = Path.Combine(_directoryWithFiles, "LicensedFont" + fsType.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + ".ttf");
            File.WriteAllBytes(fontPath, font);

            using (var document = WordDocument.Create()) {
                FontTablePart? fontTable = document._wordprocessingDocument.MainDocumentPart!.FontTablePart;
                int partsBefore = fontTable?.FontParts.Count() ?? 0;
                if (embedded) {
                    document.EmbedFont(fontPath);
                } else {
                    var exception = Assert.Throws<ArgumentException>(() => document.EmbedFont(fontPath));
                    Assert.Contains("fsType", exception.Message, StringComparison.Ordinal);
                }

                int partsAfter = document._wordprocessingDocument.MainDocumentPart!.FontTablePart?.FontParts.Count() ?? 0;
                Assert.Equal(partsBefore + (embedded ? 1 : 0), partsAfter);
            }
            File.Delete(fontPath);
        }
    }
}
