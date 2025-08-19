using System.IO;
using System.Linq;
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
            var fontTablePart = word.MainDocumentPart!.FontTablePart;
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
            var styles = word.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "DejaVuStyle"));
            File.Delete(fontPath);
        }
    }
}
