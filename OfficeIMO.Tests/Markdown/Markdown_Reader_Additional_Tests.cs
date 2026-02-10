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
        public void Blockquote_Allows_Lazy_Continuation() {
            string md = "> Quote line 1\nQuote line 2\n\nParagraph.";
            var doc = MarkdownReader.Parse(md);
            var quote = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.Single(quote.Children);
            var para = Assert.IsType<ParagraphBlock>(quote.Children[0]);
            var html = ((IMarkdownBlock)para).RenderHtml();
            Assert.Contains("Quote line 1", html);
            Assert.Contains("Quote line 2", html);
        }

        [Fact]
        public void Blockquote_Resolves_Reference_Links_From_Outer_Document() {
            string md = """
> [x][r]

[r]: https://example.com
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
            Assert.Contains("href=\"https://example.com\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Blockquote_Does_Not_Parse_Front_Matter() {
            string md = """
> ---
> title: x
> ---
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
            Assert.Contains("<dt>title</dt>", html, StringComparison.Ordinal);
            Assert.Contains("<dd>x</dd>", html, StringComparison.Ordinal);
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
        public void Parses_Angle_Bracket_Autolink_Url() {
            var doc = MarkdownReader.Parse("<https://example.com>");
            var html = doc.ToHtmlFragment();
            Assert.Contains("href=\"https://example.com\"", html);
            Assert.Contains(">https://example.com<", html);
        }

        [Fact]
        public void Parses_Angle_Bracket_Autolink_Email() {
            var doc = MarkdownReader.Parse("<user@example.com>");
            var html = doc.ToHtmlFragment();
            Assert.Contains("href=\"mailto:user@example.com\"", html);
            Assert.Contains(">user@example.com<", html);
        }

        [Fact]
        public void Parses_Link_With_Single_Quote_Title() {
            string md = "[x](https://example.com 'title')";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("href=\"https://example.com\"", html);
            Assert.Contains("title=\"title\"", html);
        }

        [Fact]
        public void Parses_Link_With_Paren_Title() {
            string md = "[x](https://example.com (title))";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("href=\"https://example.com\"", html);
            Assert.Contains("title=\"title\"", html);
        }

        [Fact]
        public void Parses_Image_With_Single_Quote_Title() {
            string md = "![alt](https://example.com/a.png 't')";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("src=\"https://example.com/a.png\"", html);
            Assert.Contains("title=\"t\"", html);
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

        [Fact]
        public void Unordered_List_Parses_Inline_Markup() {
            string md = "- *emphasis* and [link](https://example.com)\n- [x] **bold** task";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Contains(list.Items[0].Content.Items, item => item is ItalicSequenceInline);
            Assert.Contains(list.Items[0].Content.Items, item => item is LinkInline);
            Assert.Contains(list.Items[1].Content.Items, item => item is BoldSequenceInline);
        }

        [Fact]
        public void Ordered_List_Allows_Paren_Delimiter() {
            string md = "1) one\n2) two";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<OrderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
        }

        [Fact]
        public void Parses_Tilde_Fenced_Code_Block() {
            string md = "~~~csharp\nConsole.WriteLine(1);\n~~~~~~\n";
            var doc = MarkdownReader.Parse(md);
            var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
            Assert.Equal("csharp", code.Language);
            Assert.Contains("Console.WriteLine(1);", code.Content);
        }

        [Fact]
        public void Underscore_Emphasis_Does_Not_Trigger_Inside_Words() {
            var doc = MarkdownReader.Parse("foo_bar_baz");
            var html = doc.ToHtmlFragment();
            Assert.DoesNotContain("<em>", html);
            Assert.Contains("foo_bar_baz", html);
        }

        [Fact]
        public void Definition_List_Terms_Parse_Inline_Markup() {
            string md = "*Term*: Definition\n[Link](https://example.com): Another";
            var doc = MarkdownReader.Parse(md);
            var defList = Assert.IsType<DefinitionListBlock>(doc.Blocks[0]);
            var html = ((IMarkdownBlock)defList).RenderHtml();
            Assert.Contains("<em>Term</em>", html);
            Assert.Contains("href=\"https://example.com\"", html);
        }

        [Fact]
        public void Backslash_End_Of_Line_Produces_Hard_Break_In_Paragraph() {
            string md = "First\\\nSecond";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();
            Assert.Contains("First<br/>Second", html);
            Assert.DoesNotContain("First\\</p>", html);
        }

        [Fact]
        public void Double_Backslash_End_Of_Line_Keeps_One_Backslash_And_Breaks() {
            string md = "First\\\\\nSecond";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();
            Assert.Contains("First\\<br/>Second", html);
        }

        [Fact]
        public void Backslash_Hard_Breaks_Can_Be_Disabled() {
            string md = "First\\\nSecond";
            var options = new MarkdownReaderOptions { BackslashHardBreaks = false, HtmlBlocks = false, InlineHtml = false };
            var doc = MarkdownReader.Parse(md, options);
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.DoesNotContain("<br/>", html);
            Assert.Contains("First\\ Second", html);
        }

        [Fact]
        public void Two_Trailing_Spaces_Produce_Hard_Break_In_Paragraph() {
            string md = "First  \nSecond";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();
            Assert.Contains("First<br/>Second", html);
        }

        [Fact]
        public void Emphasis_Can_Nest_Inside_Bold() {
            string md = "**bold *italic* text**";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("<strong>bold <em>italic</em> text</strong>", html);
        }

        [Fact]
        public void Bold_Can_Nest_Inside_Italic() {
            string md = "*a **b** c*";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("<em>a <strong>b</strong> c</em>", html);
        }

        [Fact]
        public void Strikethrough_Can_Contain_Inline_Markup() {
            string md = "~~a **b**~~";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("<del>a <strong>b</strong></del>", html);
        }

        [Fact]
        public void CodeSpan_Trims_One_Leading_And_Trailing_Space() {
            string md = "` a `";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("<code>a</code>", html);
        }

        [Fact]
        public void CodeSpan_Trims_Only_One_Space_Per_Side() {
            string md = "`  a  `";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("<code> a </code>", html);
        }

        [Fact]
        public void Reference_Image_Full_Collapsed_And_Shortcut_Are_Supported() {
            string md = """
![logo][id]

[id]: https://example.com/logo.png "title"
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("src=\"https://example.com/logo.png\"", html);
            Assert.Contains("alt=\"logo\"", html);
            Assert.Contains("title=\"title\"", html);
        }

        [Fact]
        public void Reference_Image_Label_Normalizes_Whitespace() {
            string md = """
![logo][my   label]

[my label]: https://example.com/logo.png
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("src=\"https://example.com/logo.png\"", html);
        }

        [Fact]
        public void Reference_Definitions_Support_Single_Quote_Titles() {
            string md = """
[x][r]

[r]: https://example.com 't'
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment();
            Assert.Contains("href=\"https://example.com\"", html);
            Assert.Contains("title=\"t\"", html);
        }

        [Fact]
        public void Markdown_Output_Chooses_Title_Delimiter_To_Avoid_Escaping() {
            var doc = MarkdownDoc.Create().P(p => p.Link("x", "https://example.com", "a\"b"));
            var md = doc.ToMarkdown().Replace("\r", "");
            Assert.Contains("[x](https://example.com 'a\"b')", md);

            var doc2 = MarkdownDoc.Create().P(p => p.Link("x", "https://example.com", "a\"b'c"));
            var md2 = doc2.ToMarkdown().Replace("\r", "");
            Assert.Contains("[x](https://example.com (a\"b'c))", md2);
        }

        [Fact]
        public void Reference_Link_Definitions_Block_Script_Schemes() {
            string md = """
[x][r]

[r]: javascript:alert(1)
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.DoesNotContain("javascript:", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<a", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Footnotes_Render_As_Section_At_End_With_Backref_And_Inline_Markup() {
            string md = """
Para1[^a]

[^a]: Footnote *one*

Para2
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

            int idxPara2 = html.IndexOf("Para2", StringComparison.Ordinal);
            int idxFoot = html.IndexOf("class=\"footnotes\"", StringComparison.Ordinal);
            Assert.True(idxPara2 >= 0 && idxFoot >= 0 && idxPara2 < idxFoot);

            Assert.Contains("<sup id=\"fnref:a\"><a href=\"#fn:a\">a</a></sup>", html, StringComparison.Ordinal);
            Assert.Contains("<li id=\"fn:a\">", html, StringComparison.Ordinal);
            Assert.Contains("<em>one</em>", html, StringComparison.Ordinal);
            Assert.Contains("href=\"#fnref:a\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Footnotes_Allow_Blank_Lines_Inside_Definition_When_Indented() {
            string md = """
Text[^a]

[^a]: First

  Second
""";
            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<li id=\"fn:a\"><p>First", html, StringComparison.Ordinal);
            Assert.Contains("</p><p>Second", html, StringComparison.Ordinal);
        }
    }
}

