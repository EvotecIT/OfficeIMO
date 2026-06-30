using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Builder_Creates_Details_Block() {
            var doc = MarkdownDoc.Create()
                .Details("More info", body => body.P("Hidden text"));

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.False(details.Open);
            Assert.Equal("More info", Assert.IsType<TextRun>(details.Summary!.Inlines.Items[0]).Text);
            var child = Assert.Single(details.ChildBlocks);
            Assert.IsType<ParagraphBlock>(child);
            Assert.Equal("<details>\n<summary>More info</summary>\n\nHidden text\n</details>", ((IMarkdownBlock)details).RenderMarkdown());
            Assert.Equal("<details>\n<summary>More info</summary>\n\n<p>Hidden text</p>\n</details>", ((IMarkdownBlock)details).RenderHtml());
        }

        [Fact]
        public void Reader_RoundTrips_Details_Html() {
            string markdown = "<details open>\n<summary>  Expand  </summary>\n\nParagraph text\n</details>";

            var doc = MarkdownReader.Parse(markdown);

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.True(details.Open);
            var summaryText = Assert.IsType<TextRun>(details.Summary!.Inlines.Items[0]);
            Assert.Equal("Expand", summaryText.Text);
            Assert.Equal("<details open>", details.OpeningTag);
            Assert.Equal("</details>", details.ClosingTag);
            Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 14), details.OpeningTagSourceSpan);
            Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 10), details.ClosingTagSourceSpan);
            Assert.Equal("<summary>", details.Summary.OpeningTag);
            Assert.Equal("  Expand  ", details.Summary.SourceText);
            Assert.Equal("</summary>", details.Summary.ClosingTag);
            Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 9), details.Summary.OpeningTagSourceSpan);
            Assert.Equal(new MarkdownSourceSpan(2, 10, 2, 19), details.Summary.TextSourceSpan);
            Assert.Equal(new MarkdownSourceSpan(2, 20, 2, 29), details.Summary.ClosingTagSourceSpan);
            var child = Assert.Single(details.ChildBlocks);
            var paragraph = Assert.IsType<ParagraphBlock>(child);
            Assert.Equal("Paragraph text", paragraph.Inlines.RenderMarkdown());

            var html = ((IMarkdownBlock)details).RenderHtml();
            Assert.Equal("<details open>\n<summary>Expand</summary>\n\n<p>Paragraph text</p>\n</details>", html);
        }

        [Fact]
        public void Summary_RenderMarkdown_Preserves_Inline_Markup() {
            var summary = new SummaryBlock(new InlineSequence()
                .Text("Use ")
                .Bold("strong")
                .Text(" ")
                .Code("code"));
            var details = new DetailsBlock(summary, new[] { new ParagraphBlock(new InlineSequence().Text("Hidden text")) });

            var markdown = ((IMarkdownBlock)details).RenderMarkdown();
            var html = ((IMarkdownBlock)details).RenderHtml();

            Assert.Contains("<summary>", markdown, StringComparison.Ordinal);
            Assert.Contains("**strong**", markdown, StringComparison.Ordinal);
            Assert.Contains("`code`", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("<strong>", markdown, StringComparison.Ordinal);
            Assert.Contains("<summary>", html, StringComparison.Ordinal);
            Assert.Contains("<strong>strong</strong>", html, StringComparison.Ordinal);
            Assert.Contains("<code>code</code>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("**strong**", html, StringComparison.Ordinal);
        }
    }
}
