using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_RegisterCustomParagraphStyle() {
            var style = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
            WordParagraphStyle.RegisterCustomStyle("MyStyle", style);

            string filePath = Path.Combine(_directoryWithFiles, "CustomParagraphStyle.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Text").SetStyleId("MyStyle");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var styles = document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
                Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "MyStyle"));
            }
        }

        [Fact]
        public void Test_OverrideBuiltInParagraphStyle() {
            var original = WordParagraphStyle.GetStyleDefinition(WordParagraphStyles.Normal);
            var custom = new Style { Type = StyleValues.Paragraph, StyleId = "Normal" };
            WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, custom);

            Assert.Equal(custom, WordParagraphStyle.GetStyleDefinition(WordParagraphStyles.Normal));

            WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, original);
        }
    }
}
