using System;
using System.IO;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Html_Block_Tests {
        [Fact]
        public void Parses_Gfm_Type1_Script_Block() {
            string md = "<script>\nlet x = 1;\n</script>\n\nParagraph.";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<script>\nlet x = 1;\n</script>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type2_Html_Comment_Block() {
            string md = "<!-- start\ncontinues -->\n\nNext";

            var doc = MarkdownReader.Parse(md);

            var comment = Assert.IsType<HtmlCommentBlock>(doc.Blocks[0]);
            Assert.Equal("<!-- start\ncontinues -->", comment.Comment);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type6_Block_Until_Blank_Line() {
            string md = "<div>\n<p>inline</p>\n\nparagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div>\n<p>inline</p>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Closing_Tag_Starts_Block() {
            string md = "</div>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("</div>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type3_Processing_Instruction_Block() {
            string md = "<?xml version=\"1.0\"?>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<?xml version=\"1.0\"?>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type4_Declaration_Block() {
            string md = "<!DOCTYPE html>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<!DOCTYPE html>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type5_CData_Block() {
            string md = "<![CDATA[<p>literal</p>]]>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<![CDATA[<p>literal</p>]]>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type7_Generic_Tag_Block() {
            string md = "<span class=\"note\">\ninline\n</span>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<span class=\"note\">\ninline\n</span>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Incomplete_Block_Tag_Remains_Text() {
            string md = "<div\nParagraph";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
        }

        [Fact]
        public void Type1_RawText_Tag_Can_Start_At_Line_End() {
            string md = "<script\nalert(1);";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<script\nalert(1);", html.Html);
        }

        [Fact]
        public void Type6_Tag_With_Escaped_Quotes_Is_Recognized() {
            string md = "<div data-json=\"{\\\"key\\\":\\\"value\\\"}\">\ncontent\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div data-json=\"{\\\"key\\\":\\\"value\\\"}\">\ncontent", html.Html);
        }

        [Fact]
        public void Type6_Block_Continues_Until_Blank_Line_After_Closing_Tag() {
            string md = "<div>\n<section>\n<p>Value</p>\n</section>\n</div>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div>\n<section>\n<p>Value</p>\n</section>\n</div>\nParagraph", html.Html);
            Assert.Single(doc.Blocks);
        }

        [Fact]
        public void Type7_Tag_Requiring_Closing_Bracket_Remains_Text_When_Incomplete() {
            string md = "<span class=\"note\"\nParagraph";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
        }

        [Fact]
        public void Type6_Details_Block_Preserves_Blank_Line_Content() {
            string md = "<details>\n<summary>Summary</summary>\n\n<div>Body</div>\n</details>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.Equal("Summary", Assert.IsType<TextRun>(details.Summary!.Inlines.Items[0]).Text);
            Assert.Equal("<details>\n<summary>Summary</summary>\n\n<div>Body</div>\n</details>", ((IMarkdownBlock)details).RenderMarkdown());
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Table_Block_Preserves_Blank_Line_Separated_Sections() {
            string md = "<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>", html.Html);
        }

        [Fact]
        public void CommonMark_Profile_Type6_Table_Block_Ends_At_Blank_Line() {
            string md = "<table>\n\n  <tr>";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<table>", html.Html);
            Assert.IsType<HtmlRawBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void OfficeImo_Profile_Type6_Table_Block_Can_Preserve_Blank_Line_Content() {
            string md = "<table>\n\n  <tr>\n</table>";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(Assert.Single(doc.Blocks));
            Assert.Equal("<table>\n\n  <tr>\n</table>", html.Html);
        }

        [Fact]
        public void Type6_Closing_Tag_Does_Not_Consume_Following_Html() {
            string md = "</div>\n\n<div>\n<p>Next</p>\n</div>";

            var doc = MarkdownReader.Parse(md);

            var closingBlock = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("</div>", closingBlock.Html);

            var nextBlock = Assert.IsType<HtmlRawBlock>(doc.Blocks[1]);
            Assert.Equal("<div>\n<p>Next</p>\n</div>", nextBlock.Html);
        }

        [Fact]
        public void Type6_Details_Block_With_No_Blank_Line_After_Close_Remains_Structured() {
            string md = "<details>\n<summary>One</summary>\n\n<section>\n<p>Inner</p>\n</section>\n\n</details>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.Equal("One", Assert.IsType<TextRun>(details.Summary!.Inlines.Items[0]).Text);
            Assert.Equal("<details>\n<summary>One</summary>\n\n<section>\n<p>Inner</p>\n</section>\n\n</details>", ((IMarkdownBlock)details).RenderMarkdown());
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Details_Block_With_Nested_Table_And_Blank_Lines_Remains_Intact() {
            string md = "<details>\n<summary>Summary</summary>\n\n<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>\n\n<div>Tail</div>\n</details>";

            var doc = MarkdownReader.Parse(md);

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<summary>Summary</summary>\n\n<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>\n\n<div>Tail</div>\n</details>", ((IMarkdownBlock)details).RenderMarkdown());
        }

        [Fact]
        public void Type6_Details_Block_With_No_Blank_Line_After_Close_And_SelfClosing_Child_Remains_Structured() {
            string md = "<details>\n<summary>Summary</summary>\n<component />\n\n<div>Body</div>\n</details>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<summary>Summary</summary>\n<component />\n\n<div>Body</div>\n</details>", ((IMarkdownBlock)details).RenderMarkdown());
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }


        [Theory]
        [InlineData("html-comment-single.md")]
        [InlineData("html-comment-multi.md")]
        public void Html_Comment_Fixtures_Parse(string fixtureName) {
            string markdown = LoadFixture(fixtureName);
            var doc = MarkdownReader.Parse(markdown);
            Assert.Single(doc.Blocks);
            var comment = Assert.IsType<HtmlCommentBlock>(doc.Blocks[0]);
            string expected = NormalizeFixture(markdown);
            Assert.Equal(expected, comment.Comment);
        }

        private static string LoadFixture(string name) {
            var path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Markdown", "Fixtures", name);
            path = Path.GetFullPath(path);
            return File.ReadAllText(path);
        }

        private static string NormalizeFixture(string content) {
            string normalized = content.Replace("\r\n", "\n").Replace('\r', '\n');
            return normalized.TrimEnd('\n');
        }
        [Fact]
        public void Type6_Details_Block_With_No_Blank_Line_After_Close_And_Unmatched_Inner_Tags_Remains_Structured() {
            string md = "<details>\n<div>\n<p>Loose</p>\n</details>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var details = Assert.IsType<DetailsBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<div>\n<p>Loose</p>\n</details>", ((IMarkdownBlock)details).RenderMarkdown());
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Html_Blocks_Can_Be_Disabled() {
            string md = "<div>Inline</div>\n\nParagraph";

            var options = new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false };
            var doc = MarkdownReader.Parse(md, options);

            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
            var firstParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            var firstText = Assert.IsType<TextRun>(firstParagraph.Inlines.Items[0]);
            Assert.Contains("<div>Inline", firstText.Text);

            bool hasClosingTag = false;
            foreach (var inline in firstParagraph.Inlines.Items) {
                if (inline is TextRun run && run.Text.IndexOf("</div>", StringComparison.Ordinal) >= 0) { hasClosingTag = true; break; }
            }
            Assert.True(hasClosingTag, "Closing tag should remain in the paragraph text.");

            var secondParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
            var secondText = Assert.IsType<TextRun>(secondParagraph.Inlines.Items[0]);
            Assert.Equal("Paragraph", secondText.Text);
        }

        [Fact]
        public void CommonMark_Profile_Preserves_Inline_Raw_Html_Constructs() {
            string md = "foo <!-- this is a --\ncomment - with hyphens -->\n\nfoo <?php echo $a; ?>\n\nfoo <!ELEMENT br EMPTY>\n\nfoo <![CDATA[>&<]]>";

            string html = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile())
                .ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            const string expected = "<p>foo <!-- this is a --\ncomment - with hyphens --></p>\n<p>foo <?php echo $a; ?></p>\n<p>foo <!ELEMENT br EMPTY></p>\n<p>foo <![CDATA[>&<]]></p>\n";
            Assert.Equal(CommonMarkHtmlComparison.Normalize(expected), CommonMarkHtmlComparison.Normalize(html));
        }

        [Fact]
        public void CommonMark_Profile_Rejects_Malformed_Inline_Raw_Html_Tags() {
            string md = "<a h*#ref=\"hi\">\n\n<a href='bar'title=title>\n\n<bar/ >\n\n<foo bar=baz\nbim!bop />";

            string html = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile())
                .ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            const string expected = "<p>&lt;a h*#ref=&quot;hi&quot;&gt;</p>\n<p>&lt;a href='bar'title=title&gt;</p>\n<p>&lt;bar/ &gt;</p>\n<p>&lt;foo bar=baz\nbim!bop /&gt;</p>\n";
            Assert.Equal(CommonMarkHtmlComparison.Normalize(expected), CommonMarkHtmlComparison.Normalize(html));
        }

        [Fact]
        public void OfficeIMO_Profile_Rejects_Malformed_Type6_Block_Html_Starts() {
            string md = "<div class\n# Heading";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateOfficeIMOProfile());

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.Equal("<div class", paragraph.Inlines.RenderMarkdown());
            Assert.IsType<HeadingBlock>(doc.Blocks[1]);
            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
        }

        [Fact]
        public void CommonMark_Profile_Treats_Comment_Shorthand_As_Raw_Html_Then_Text() {
            string md = "foo <!--> foo -->\n\nfoo <!---> foo -->";

            string html = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile())
                .ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            const string expected = "<p>foo <!--> foo --&gt;</p>\n<p>foo <!---> foo --&gt;</p>\n";
            Assert.Equal(CommonMarkHtmlComparison.Normalize(expected), CommonMarkHtmlComparison.Normalize(html));
        }

        [Fact]
        public void CommonMark_Profile_Renders_Blockquote_Raw_Html_Block_Boundary() {
            string md = "> <div>\n> foo\n\nbar\n";

            string html = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile())
                .ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            const string expected = "<blockquote>\n<div>\nfoo\n</blockquote>\n<p>bar</p>\n";
            Assert.Equal(CommonMarkHtmlComparison.Normalize(expected), CommonMarkHtmlComparison.Normalize(html));
        }
    }
}

