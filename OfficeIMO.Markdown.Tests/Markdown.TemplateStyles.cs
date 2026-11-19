using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class MarkdownTemplateStyles {
        [Fact]
        public void MarkdownIntoTemplate_WithDocumentStyles_UsesTemplateHeadingBodyAndQuoteStyles() {
            using var stream = new MemoryStream();
            using (WordDocument template = WordDocument.Create(stream)) {
                template.Styles.AddParagraphStyle("CorpsTexte", "Corps de texte").FontFamily = "Garamond";
                template.Styles.Remove("Heading1");
                template.Styles.AddParagraphStyle("Titre1", "heading 1").Bold = true;
                template.Styles.AddParagraphStyle("Citation", "Quote").Italic = true;
                template.AddParagraph("PLACEHOLDER").AddBookmark("Content");
                template.Save();
            }
            stream.Position = 0;
            using WordDocument document = WordDocument.Load(stream);
            const string markdown = """
                # Report

                Body text.

                > Quoted
                >
                > > Nested

                - Item
                """;
            var options = new MarkdownWordExportOptions {
                Template = document,
                BookmarkName = "Content",
                PreferDocumentStyles = true,
                DefaultParagraphStyleId = "Corps de texte"
            };

            OfficeIMO.Markdown.MarkdownDocument.Parse(markdown, options.CreateReaderOptions()).ToWord(options);

            List<Paragraph> paragraphs = document.BodyRoot.Elements<Paragraph>().ToList();
            Assert.Equal("Titre1", StyleOf(paragraphs, "Report"));
            Assert.Equal("CorpsTexte", StyleOf(paragraphs, "Body text."));
            Paragraph quoted = Assert.Single(paragraphs, paragraph => paragraph.InnerText == "Quoted");
            Assert.Equal("Citation", quoted.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.Null(quoted.ParagraphProperties?.Indentation);
            Paragraph nested = Assert.Single(paragraphs, paragraph => paragraph.InnerText == "Nested");
            Assert.Equal("Citation", nested.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.Equal("720", nested.ParagraphProperties?.Indentation?.Left?.Value);
            Assert.NotEqual("CorpsTexte", StyleOf(paragraphs, "Item"));
            Assert.False(document.Styles.Contains("Heading1"));
        }

        [Fact]
        public void MarkdownToWord_WithUnknownDefaultParagraphStyle_Throws() {
            var options = new MarkdownWordExportOptions { DefaultParagraphStyleId = "MissingBody" };

            var exception = Assert.Throws<InvalidOperationException>(() =>
                OfficeIMO.Markdown.MarkdownDocument.Parse("Body").ToWord(options));

            Assert.Contains("MissingBody", exception.Message, StringComparison.Ordinal);
        }

        private static string? StyleOf(IEnumerable<Paragraph> paragraphs, string text) =>
            Assert.Single(paragraphs, paragraph => paragraph.InnerText.Trim() == text).ParagraphProperties?.ParagraphStyleId?.Val?.Value;
    }
}
