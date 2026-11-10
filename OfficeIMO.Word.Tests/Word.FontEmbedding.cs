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

                Assert.False(document.Styles.Contains("BrandFont"));
                Assert.Equal(fontsBefore, document._wordprocessingDocument.MainDocumentPart!.FontTablePart?.Fonts?.Elements<Font>().Count() ?? 0);
            }
            File.Delete(fontPath);
        }
    }
}
