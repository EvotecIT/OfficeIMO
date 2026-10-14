using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(WordFileFormat.Docx)]
        [InlineData(WordFileFormat.Doc)]
        public void ClearingParagraphAlignmentRestoresStyleInheritance(WordFileFormat format) {
            using WordDocument document = WordDocument.Create();
            const string styleId = "InheritedParagraphAlignment";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new Style(new StyleName { Val = styleId }, new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new Justification { Val = JustificationValues.Center },
                        new TextAlignment { Val = VerticalTextAlignmentValues.Center })) {
                    Type = StyleValues.Paragraph, StyleId = styleId, CustomStyle = true
                });
            WordParagraph paragraph = document.AddParagraph("Inherited alignment");
            paragraph._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
            paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
            paragraph.VerticalCharacterAlignmentOnLine = WordVerticalCharacterAlignment.Bottom;

            paragraph.ParagraphAlignment = null;
            paragraph.VerticalCharacterAlignmentOnLine = null;
            byte[] bytes = document.ToBytes(format);
            if (format == WordFileFormat.Doc) {
                using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
                bytes = loaded.ToBytes();
            }

            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            ParagraphProperties properties = package.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single().ParagraphProperties!;
            Assert.Null(properties.Justification);
            Assert.Null(properties.TextAlignment);
            string? importedStyleId = properties.ParagraphStyleId?.Val?.Value;
            Assert.False(string.IsNullOrEmpty(importedStyleId));
            if (format == WordFileFormat.Docx) Assert.Equal(styleId, importedStyleId);
            Style inheritedStyle = Assert.Single(package.MainDocumentPart.StyleDefinitionsPart!.Styles!.Elements<Style>(),
                item => item.StyleId?.Value == importedStyleId);
            Assert.Equal(styleId, inheritedStyle.StyleName?.Val?.Value);
            Assert.Equal(JustificationValues.Center, inheritedStyle.StyleParagraphProperties?.GetFirstChild<Justification>()?.Val?.Value);
            Assert.Equal(VerticalTextAlignmentValues.Center, inheritedStyle.StyleParagraphProperties?.GetFirstChild<TextAlignment>()?.Val?.Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LegacyDoc_UsesNativeParagraphSpacingDefaults(bool builtInStyle) {
            byte[] bytes = builtInStyle
                ? LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleFontFamily()
                : LegacyDocTestBuilder.CreateSimpleDoc("First paragraph", "Second paragraph");
            using WordDocument document = WordDocument.Load(new MemoryStream(bytes));
            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);
            Styles styles = package.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            SpacingBetweenLines spacing = styles.DocDefaults!.ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!.GetFirstChild<SpacingBetweenLines>()!;

            Assert.Equal("240", spacing.Line?.Value);
            Assert.Equal(LineSpacingRuleValues.Auto, spacing.LineRule?.Value);
            Assert.Equal("0", spacing.Before?.Value);
            Assert.Equal("0", spacing.After?.Value);
            if (builtInStyle) {
                Style heading = Assert.Single(styles.Elements<Style>(), item => item.StyleId?.Value == "Heading1");
                Assert.Null(heading.StyleParagraphProperties?.GetFirstChild<SpacingBetweenLines>());
            }
        }
    }
}
