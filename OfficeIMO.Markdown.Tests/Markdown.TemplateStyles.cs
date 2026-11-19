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
                WordStyle quoteStyle = template.Styles.AddParagraphStyle("Citation", "Quote");
                quoteStyle.Italic = true;
                quoteStyle.IndentationBefore = OfficeLength.Points(48);
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

                <details>
                <summary>More</summary>

                Hidden detail

                </details>
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
            // A direct indent replaces the quote style's 48pt, so the nested level adds 36pt to it.
            Assert.Equal("1680", nested.ParagraphProperties?.Indentation?.Left?.Value);
            Assert.NotEqual("CorpsTexte", StyleOf(paragraphs, "Item"));
            Assert.False(document.Styles.Contains("Heading1"));

            Assert.Equal("CorpsTexte", StyleOf(paragraphs, "Hidden detail"));
            string exported = document.ToMarkdown().ToMarkdown();
            Assert.Contains("> Quoted", exported, StringComparison.Ordinal);
            Assert.Contains("> > Nested", exported, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownIntoTemplate_QuoteDepthRoundTripsThroughAnyQuoteStyle() {
            using var stream = new MemoryStream();
            using (WordDocument template = WordDocument.Create(stream)) {
                template.Styles.AddParagraphStyle("Quote", "Quote");
                Styles styles = template._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.DocDefaults!.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle());
                styles.DocDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle ??= new ParagraphPropertiesBaseStyle();
                styles.DocDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle.Indentation = new Indentation { Left = "1440" };
                template.AddParagraph("PLACEHOLDER").AddBookmark("Content");
                template.Save();
            }
            stream.Position = 0;
            using WordDocument document = WordDocument.Load(stream);
            var options = new MarkdownWordExportOptions { Template = document, BookmarkName = "Content", DefaultParagraphStyleId = "Quote" };

            OfficeIMO.Markdown.MarkdownDocument.Parse("> Outer\n>\n> > Inner", options.CreateReaderOptions()).ToWord(options);

            Paragraph outer = Assert.Single(document.BodyRoot.Elements<Paragraph>(), paragraph => paragraph.InnerText == "Outer");
            Paragraph inner = Assert.Single(document.BodyRoot.Elements<Paragraph>(), paragraph => paragraph.InnerText == "Inner");
            Assert.Null(outer.ParagraphProperties?.Indentation);
            // The quote style inherits the 1440-twip default indent, and the nested level adds 720 to it.
            Assert.Equal("2160", inner.ParagraphProperties?.Indentation?.Left?.Value);
            string exported = document.ToMarkdown().ToMarkdown();
            Assert.Contains("> Outer", exported, StringComparison.Ordinal);
            Assert.DoesNotContain("> > Outer", exported, StringComparison.Ordinal);
            Assert.Contains("> > Inner", exported, StringComparison.Ordinal);
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
