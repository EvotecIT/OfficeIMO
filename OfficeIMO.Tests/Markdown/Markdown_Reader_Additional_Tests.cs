using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Additional_Tests {
        [Fact]
        public void Parses_Blockquote_And_Hr() {
            string md = "> Quote line 1\n> Quote line 2\n\n---\n\nParagraph.";
            var doc = MarkdownReader.Parse(md);
            Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.IsType<HorizontalRuleBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Autolink_And_HtmlBlock() {
            string md = "Check https://example.com.\n\n<div>hi</div>\n<p>raw</p>";
            var doc = MarkdownReader.Parse(md);
            // Expect paragraph, then html block
            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.IsType<HtmlRawBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Inline_Html_Br_Can_Be_Disabled() {
            string md = "First<br>Second";

            var options = new MarkdownReaderOptions { InlineHtml = false, HtmlBlocks = false };
            var doc = MarkdownReader.Parse(md, options);

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.Single(paragraph.Inlines.Items);
            var text = Assert.IsType<TextRun>(paragraph.Inlines.Items[0]);
            Assert.Equal("First<br>Second", text.Text);
        }

        [Fact]
        public void Inline_Html_Underline_Can_Be_Disabled() {
            string md = "<u>Decorated</u>";

            var options = new MarkdownReaderOptions { InlineHtml = false, HtmlBlocks = false };
            var doc = MarkdownReader.Parse(md, options);

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.Single(paragraph.Inlines.Items);
            var text = Assert.IsType<TextRun>(paragraph.Inlines.Items[0]);
            Assert.Equal("<u>Decorated</u>", text.Text);
        }

        [Fact]
        public void Inline_Html_Remains_When_Html_Blocks_Disabled() {
            string md = "<div>First<br>Second</div>";

            var options = new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = true };
            var doc = MarkdownReader.Parse(md, options);

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.Equal(4, paragraph.Inlines.Items.Count);
            var firstText = Assert.IsType<TextRun>(paragraph.Inlines.Items[0]);
            Assert.Equal("<div>First", firstText.Text);
            Assert.IsType<HardBreakInline>(paragraph.Inlines.Items[1]);
            var secondText = Assert.IsType<TextRun>(paragraph.Inlines.Items[2]);
            Assert.Equal("Second", secondText.Text);
            var closingTag = Assert.IsType<TextRun>(paragraph.Inlines.Items[3]);
            Assert.Equal("</div>", closingTag.Text);
        }

        [Fact]
        public void Html_Blocks_Remain_When_Inline_Html_Disabled() {
            string md = "<div>Inline <br/> html</div>\n\nParagraph";

            var options = new MarkdownReaderOptions { HtmlBlocks = true, InlineHtml = false };
            var doc = MarkdownReader.Parse(md, options);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div>Inline <br/> html</div>", html.Html);
            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
            var text = Assert.IsType<TextRun>(paragraph.Inlines.Items[0]);
            Assert.Equal("Paragraph", text.Text);
        }

        [Fact]
        public void Heading_With_Colon_Is_Not_Definition_List() {
            string md = "## Heading: Text\n\nParagraph.";
            var doc = MarkdownReader.Parse(md);
            var heading = Assert.IsType<HeadingBlock>(doc.Blocks[0]);
            Assert.Equal(2, heading.Level);
            Assert.Equal("Heading: Text", heading.Text);
        }
    }
}

