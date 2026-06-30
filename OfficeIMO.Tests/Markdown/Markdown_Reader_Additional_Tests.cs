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
        public void Blockquote_Lazy_Continuation_Does_Not_Swallow_Following_List() {
            string md = "> Quote line 1\n- outside item";
            var doc = MarkdownReader.Parse(md);

            Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.IsType<UnorderedListBlock>(doc.Blocks[1]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<blockquote><p>Quote line 1</p></blockquote>", html, StringComparison.Ordinal);
            Assert.Contains("<ul><li>outside item</li></ul>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Blockquote_Lazy_Continuation_Keeps_Indented_Text_Inside_Quote() {
            string md = "> Quote line 1\n    outside code";
            var doc = MarkdownReader.Parse(md);

            var quote = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.Single(doc.Blocks);
            var paragraph = Assert.IsType<ParagraphBlock>(quote.Children[0]);
            var html = ((IMarkdownBlock)paragraph).RenderHtml();
            Assert.Contains("Quote line 1\noutside code", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Quote line 1  outside code", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Blockquote_Lazy_Continuation_Keeps_Indented_List_Like_Text_Inside_Quote() {
            string md = "> Quote line 1\n    - nested";
            var doc = MarkdownReader.Parse(md);

            var quote = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.Single(doc.Blocks);
            Assert.Single(quote.Children);
            var paragraph = Assert.IsType<ParagraphBlock>(quote.Children[0]);
            var html = ((IMarkdownBlock)paragraph).RenderHtml();
            Assert.Contains("Quote line 1\n- nested", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<ul>", html, StringComparison.Ordinal);
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
        public void Parses_Standalone_Image_With_Balanced_Parens_In_Destination() {
            var doc = MarkdownReader.Parse("![alt](https://example.com/a_(b).png)");

            var image = Assert.IsType<ImageBlock>(doc.Blocks[0]);
            Assert.Equal("alt", image.Alt);
            Assert.Equal("https://example.com/a_(b).png", image.Path);
        }

        [Fact]
        public void Parses_Standalone_Image_With_Angle_Destination_Containing_Spaces() {
            var doc = MarkdownReader.Parse("![alt](<https://example.com/a (b).png>)");

            var image = Assert.IsType<ImageBlock>(doc.Blocks[0]);
            Assert.Equal("alt", image.Alt);
            Assert.Equal("https://example.com/a (b).png", image.Path);
        }

        [Fact]
        public void Parses_Linked_Image_With_Caption_As_ImageBlock() {
            var doc = MarkdownReader.Parse("""
[![alt](https://example.com/a.png)](https://example.com/docs "Documentation")
_Caption_
""");

            var image = Assert.IsType<ImageBlock>(doc.Blocks[0]);
            Assert.Equal("alt", image.Alt);
            Assert.Equal("https://example.com/a.png", image.Path);
            Assert.Equal("https://example.com/docs", image.LinkUrl);
            Assert.Equal("Documentation", image.LinkTitle);
            Assert.Equal("Caption", image.Caption);
        }

        [Fact]
        public void Parses_Inline_Image_With_Angle_Destination_Containing_Spaces() {
            var html = MarkdownReader.Parse("Look ![alt](<https://example.com/a (b).png>) now").ToHtmlFragment();

            Assert.Contains("src=\"https://example.com/a%20(b).png\"", html);
            Assert.Contains("alt=\"alt\"", html);
        }

        [Fact]
        public void Parses_Inline_Link_With_Angle_Destination_Containing_Spaces() {
            var html = MarkdownReader.Parse("[x](<https://example.com/a b> \"title\")").ToHtmlFragment();

            Assert.Contains("href=\"https://example.com/a%20b\"", html, StringComparison.Ordinal);
            Assert.Contains("title=\"title\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Reference_Link_With_Angle_Destination_Containing_Spaces_Is_Preserved() {
            const string md = """
[x][r]

[r]: <https://example.com/a b>
""";

            var doc = MarkdownReader.Parse(md);
            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            var link = Assert.IsType<LinkInline>(paragraph.Inlines.Items[0]);

            Assert.Equal("https://example.com/a b", link.Url);

            var html = doc.ToHtmlFragment();
            Assert.Contains("href=\"https://example.com/a%20b\"", html, StringComparison.Ordinal);
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
            Assert.Equal(5, paragraph.Inlines.Items.Count);
            var openingTag = Assert.IsType<HtmlRawInline>(paragraph.Inlines.Items[0]);
            Assert.Equal("<div>", openingTag.Html);
            var firstText = Assert.IsType<TextRun>(paragraph.Inlines.Items[1]);
            Assert.Equal("First", firstText.Text);
            Assert.IsType<HardBreakInline>(paragraph.Inlines.Items[2]);
            var secondText = Assert.IsType<TextRun>(paragraph.Inlines.Items[3]);
            Assert.Equal("Second", secondText.Text);
            var closingTag = Assert.IsType<HtmlRawInline>(paragraph.Inlines.Items[4]);
            Assert.Equal("</div>", closingTag.Html);
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
        public void Atx_Heading_Parses_Inline_Markup_And_Uses_PlainText_For_Slug() {
            const string md = "## **Heading** `Text`";

            var doc = MarkdownReader.Parse(md);

            var heading = Assert.IsType<HeadingBlock>(doc.Blocks[0]);
            Assert.Equal("Heading Text", heading.Text);
            Assert.Equal("## **Heading** `Text`", ((IMarkdownBlock)heading).RenderMarkdown());

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<h2 id=\"heading-text\"><strong>Heading</strong> <code>Text</code></h2>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Atx_Heading_Plain_Text_Uses_Link_Label_Text_From_Inline_Contracts() {
            const string md = "## Prefix [Linked `Text`](https://example.com)";

            var doc = MarkdownReader.Parse(md);

            var heading = Assert.IsType<HeadingBlock>(doc.Blocks[0]);
            Assert.Equal("Prefix Linked Text", heading.Text);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("id=\"prefix-linked-text\"", html, StringComparison.Ordinal);
            Assert.Contains("<a href=\"https://example.com\">Linked <code>Text</code></a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Setext_Heading_Parses_Inline_Markup() {
            const string md = """
                **Heading** `Text`
                ------------------
                """;

            var doc = MarkdownReader.Parse(md);

            var heading = Assert.IsType<HeadingBlock>(doc.Blocks[0]);
            Assert.Equal(2, heading.Level);
            Assert.Equal("Heading Text", heading.Text);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<h2 id=\"heading-text\"><strong>Heading</strong> <code>Text</code></h2>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Invalid_Reference_Definition_Like_Line_Does_Not_Become_Definition_List() {
            const string md = "[a [b]]: https://example.com";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.DoesNotContain("<dl>", html, StringComparison.Ordinal);
            Assert.Contains("<p>", html, StringComparison.Ordinal);
            Assert.Contains("[a [b]]:", html, StringComparison.Ordinal);
            Assert.Contains("https://example.com", html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void PreferNarrativeSingleLineDefinitions_Keeps_Isolated_Label_Value_As_Paragraph() {
            const string md = "Interpretation: topology looks clean in this sample.";

            var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions {
                PreferNarrativeSingleLineDefinitions = true
            });

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain("<dl>", doc.ToHtmlFragment(), StringComparison.Ordinal);
        }

        [Fact]
        public void PreferNarrativeSingleLineDefinitions_Still_Parses_Grouped_Definition_List() {
            const string md = """
                Status: healthy
                Impact: none
                """;

            var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions {
                PreferNarrativeSingleLineDefinitions = true
            });

            var definitionList = Assert.IsType<DefinitionListBlock>(doc.Blocks[0]);
            Assert.Equal(2, definitionList.Items.Count);
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
        public void Parses_Fenced_Code_Block_InfoString_Into_Primary_Language_And_Metadata() {
            const string md = """
```json title="summary chart"
{"value":1}
```
""";

            var doc = MarkdownReader.Parse(md);
            var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);

            Assert.Equal("json", code.Language);
            Assert.Equal("json title=\"summary chart\"", code.InfoString);
            Assert.Equal("title=\"summary chart\"", code.FenceInfo.AdditionalInfo);
            Assert.True(code.FenceInfo.TryGetAttribute("title", out var title));
            Assert.Equal("summary chart", title);
            Assert.Equal("summary chart", code.FenceInfo.Title);
            Assert.Equal(
                md.TrimEnd().Replace("\r\n", "\n"),
                ((IMarkdownBlock)code).RenderMarkdown().Replace("\r\n", "\n"));
        }

        [Fact]
        public void Parses_Fenced_Code_Block_Brace_Metadata_Into_Id_Classes_And_Attributes() {
            const string md = """
```chart {#quarterly-overview .wide .accent title="Quarterly Revenue" pinned}
{"series":[1,2,3]}
```
""";

            var result = MarkdownReader.ParseWithSyntaxTree(md);
            var code = Assert.IsType<CodeBlock>(result.Document.Blocks[0]);

            Assert.Equal("chart", code.Language);
            Assert.Equal("{#quarterly-overview .wide .accent title=\"Quarterly Revenue\" pinned}", code.FenceInfo.AdditionalInfo);
            Assert.Equal("quarterly-overview", code.FenceInfo.ElementId);
            Assert.Equal(2, code.FenceInfo.Classes.Count);
            Assert.Equal("wide", code.FenceInfo.Classes[0]);
            Assert.Equal("accent", code.FenceInfo.Classes[1]);
            Assert.True(code.FenceInfo.HasClass("wide"));
            Assert.True(code.FenceInfo.HasClass("accent"));
            Assert.True(code.FenceInfo.TryGetAttribute("pinned", out var pinned));
            Assert.Equal("true", pinned);
            Assert.Equal("Quarterly Revenue", code.FenceInfo.Title);
            Assert.Equal("quarterly-overview", code.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, code.Attributes.Classes);
            Assert.True(code.Attributes.TryGetAttribute("pinned", out var genericPinned));
            Assert.Equal("true", genericPinned);
            Assert.Equal("Quarterly Revenue", code.Attributes.GetAttribute("title"));

            var syntax = Assert.Single(result.SyntaxTree.Children);
            Assert.Equal("quarterly-overview", syntax.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, syntax.Attributes.Classes);
            Assert.Equal("Quarterly Revenue", syntax.Attributes.GetAttribute("title"));

            var finalSyntax = Assert.Single(result.FinalSyntaxTree.Children);
            Assert.Equal("quarterly-overview", finalSyntax.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, finalSyntax.Attributes.Classes);
            Assert.Equal("true", finalSyntax.Attributes.GetAttribute("pinned"));
            Assert.Equal(
                md.TrimEnd().Replace("\r\n", "\n"),
                ((IMarkdownBlock)code).RenderMarkdown().Replace("\r\n", "\n"));

            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            Assert.Contains("<pre id=\"quarterly-overview\" class=\"wide accent\" pinned=\"true\" title=\"Quarterly Revenue\"><code class=\"language-chart\">", html, StringComparison.Ordinal);
        }

        [Fact]
        public void GenericAttributes_Parse_AtxHeading_TrailingAttributeBlock_WhenEnabled() {
            const string md = "# Quarterly Revenue {#quarterly-overview .wide .accent title=\"Quarterly Revenue\" pinned}";
            var options = new MarkdownReaderOptions {
                GenericAttributes = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(md, options);
            var heading = Assert.IsType<HeadingBlock>(result.Document.Blocks[0]);

            Assert.Equal("Quarterly Revenue", heading.Text);
            Assert.Equal("quarterly-overview", heading.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, heading.Attributes.Classes);
            Assert.Equal("Quarterly Revenue", heading.Attributes.GetAttribute("title"));
            Assert.Equal("true", heading.Attributes.GetAttribute("pinned"));

            var syntax = Assert.Single(result.SyntaxTree.Children);
            Assert.Equal("quarterly-overview", syntax.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, syntax.Attributes.Classes);
            Assert.Equal("true", syntax.Attributes.GetAttribute("pinned"));

            var finalSyntax = Assert.Single(result.FinalSyntaxTree.Children);
            Assert.Equal("quarterly-overview", finalSyntax.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, finalSyntax.Attributes.Classes);

            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            Assert.Contains("<h1 id=\"quarterly-overview\" class=\"wide accent\" pinned=\"true\" title=\"Quarterly Revenue\">Quarterly Revenue</h1>", html, StringComparison.Ordinal);

            var written = result.Document.ToMarkdown().TrimEnd().Replace("\r\n", "\n");
            Assert.Equal("# Quarterly Revenue {#quarterly-overview .wide .accent pinned title=\"Quarterly Revenue\"}", written);

            var reparsed = MarkdownReader.Parse(written, options);
            var reparsedHeading = Assert.IsType<HeadingBlock>(reparsed.Blocks[0]);
            Assert.Equal("quarterly-overview", reparsedHeading.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, reparsedHeading.Attributes.Classes);
            Assert.Equal("Quarterly Revenue", reparsedHeading.Attributes.GetAttribute("title"));
        }

        [Fact]
        public void GenericAttributes_Parse_SetextHeading_TrailingAttributeBlock_WhenEnabled() {
            const string md = """
Quarterly Revenue {#quarterly-overview .wide .accent title="Quarterly Revenue" pinned}
=================
""";
            var options = new MarkdownReaderOptions {
                GenericAttributes = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(md, options);
            var heading = Assert.IsType<HeadingBlock>(result.Document.Blocks[0]);

            Assert.Equal(1, heading.Level);
            Assert.Equal("Quarterly Revenue", heading.Text);
            Assert.Equal("quarterly-overview", heading.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, heading.Attributes.Classes);
            Assert.Equal("Quarterly Revenue", heading.Attributes.GetAttribute("title"));
            Assert.Equal("true", heading.Attributes.GetAttribute("pinned"));

            var syntax = Assert.Single(result.SyntaxTree.Children);
            Assert.Equal("quarterly-overview", syntax.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, syntax.Attributes.Classes);
            Assert.Equal("true", syntax.Attributes.GetAttribute("pinned"));

            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            Assert.Contains("<h1 id=\"quarterly-overview\" class=\"wide accent\" pinned=\"true\" title=\"Quarterly Revenue\">Quarterly Revenue</h1>", html, StringComparison.Ordinal);

            var written = result.Document.ToMarkdown().TrimEnd().Replace("\r\n", "\n");
            Assert.Equal("# Quarterly Revenue {#quarterly-overview .wide .accent pinned title=\"Quarterly Revenue\"}", written);

            var reparsed = MarkdownReader.Parse(written, options);
            var reparsedHeading = Assert.IsType<HeadingBlock>(reparsed.Blocks[0]);
            Assert.Equal("quarterly-overview", reparsedHeading.Attributes.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, reparsedHeading.Attributes.Classes);
            Assert.Equal("Quarterly Revenue", reparsedHeading.Attributes.GetAttribute("title"));
        }

        [Fact]
        public void GenericAttributes_Parse_Paragraph_TrailingAttributeBlock_WhenEnabled() {
            const string md = "Lead paragraph {#lead .intro data-kind=\"summary\" pinned}";
            var options = new MarkdownReaderOptions {
                GenericAttributes = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(md, options);
            var paragraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[0]);

            Assert.Equal("Lead paragraph", paragraph.Inlines.RenderMarkdown());
            Assert.Equal("lead", paragraph.Attributes.ElementId);
            Assert.Equal(new[] { "intro" }, paragraph.Attributes.Classes);
            Assert.Equal("summary", paragraph.Attributes.GetAttribute("data-kind"));
            Assert.Equal("true", paragraph.Attributes.GetAttribute("pinned"));
            Assert.Equal(" ", paragraph.GenericAttributeConsumedWhitespace);

            var syntax = Assert.Single(result.SyntaxTree.Children);
            Assert.Equal("lead", syntax.Attributes.ElementId);
            Assert.Equal(new[] { "intro" }, syntax.Attributes.Classes);
            Assert.Equal("summary", syntax.Attributes.GetAttribute("data-kind"));

            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            Assert.Contains("<p id=\"lead\" class=\"intro\" data-kind=\"summary\" pinned=\"true\">Lead paragraph </p>", html, StringComparison.Ordinal);

            var written = result.Document.ToMarkdown().TrimEnd().Replace("\r\n", "\n");
            Assert.Equal("Lead paragraph {#lead .intro data-kind=\"summary\" pinned}", written);

            var reparsed = MarkdownReader.Parse(written, options);
            var reparsedParagraph = Assert.IsType<ParagraphBlock>(reparsed.Blocks[0]);
            Assert.Equal("lead", reparsedParagraph.Attributes.ElementId);
            Assert.Equal(new[] { "intro" }, reparsedParagraph.Attributes.Classes);
            Assert.Equal("summary", reparsedParagraph.Attributes.GetAttribute("data-kind"));
        }

        [Fact]
        public void GenericAttributes_Parse_PipeTable_CellAttributeBlock_As_TableAttributes_WhenEnabled() {
            const string md = """
| A {#tbl .wide title="Quarterly"} |
|---|
| B |
""";
            var options = new MarkdownReaderOptions {
                GenericAttributes = true,
                PreserveTrivia = true,
                Tables = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(md, options);
            var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Blocks));

            Assert.Equal("tbl", table.Attributes.ElementId);
            Assert.Equal(new[] { "wide" }, table.Attributes.Classes);
            Assert.Equal("Quarterly", table.Attributes.GetAttribute("title"));
            Assert.Equal("A", Assert.Single(table.Headers));
            Assert.Equal("A", Assert.Single(table.HeaderCells).Markdown);
            Assert.Equal("B", Assert.Single(Assert.Single(table.Rows)));

            var syntax = Assert.Single(result.FinalSyntaxTree.Children);
            Assert.Equal(MarkdownSyntaxKind.Table, syntax.Kind);
            Assert.Equal("tbl", syntax.Attributes.ElementId);
            var attributes = Assert.Single(syntax.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
            Assert.Equal("{#tbl .wide title=\"Quarterly\"}", attributes.Literal);
            Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
            Assert.Equal("{#tbl .wide title=\"Quarterly\"}", slice.Text);

            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            Assert.Contains("<table id=\"tbl\" class=\"wide\" title=\"Quarterly\">", html, StringComparison.Ordinal);
            Assert.Contains("<th>A</th>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("{#tbl", html, StringComparison.Ordinal);

            var written = result.Document.ToMarkdown().TrimEnd().Replace("\r\n", "\n");
            Assert.Equal("| A {#tbl .wide title=\"Quarterly\"} |\n| --- |\n| B |", written);

            var reparsed = MarkdownReader.Parse(written, options);
            var reparsedTable = Assert.IsType<TableBlock>(Assert.Single(reparsed.Blocks));
            Assert.Equal("tbl", reparsedTable.Attributes.ElementId);
            Assert.Equal("A", Assert.Single(reparsedTable.Headers));
        }

        [Fact]
        public void GenericAttributes_Parse_InlineElements_WhenEnabled() {
            const string md = "[site](https://example.com){#lnk .primary title=\"Site\"} ![alt](image.png){#img .wide title=\"Image\"} *emphasis*{#em .marked} **strong**{#strong .marked} `code`{#code .token}";
            var options = new MarkdownReaderOptions {
                GenericAttributes = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(md, options);
            var paragraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[0]);

            AssertAttributes(Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>()).Attributes, "lnk", new[] { "primary" }, ("title", "Site"));
            AssertAttributes(Assert.Single(paragraph.Inlines.Nodes.OfType<ImageInline>()).Attributes, "img", new[] { "wide" }, ("title", "Image"));
            AssertAttributes(Assert.Single(paragraph.Inlines.Nodes.OfType<ItalicSequenceInline>()).Attributes, "em", new[] { "marked" });
            AssertAttributes(Assert.Single(paragraph.Inlines.Nodes.OfType<BoldSequenceInline>()).Attributes, "strong", new[] { "marked" });
            AssertAttributes(Assert.Single(paragraph.Inlines.Nodes.OfType<CodeSpanInline>()).Attributes, "code", new[] { "token" });
            Assert.DoesNotContain(paragraph.Inlines.Nodes.OfType<TextRun>(), text => text.Text.Contains("{#", StringComparison.Ordinal));

            var syntax = Assert.Single(result.SyntaxTree.Children);
            AssertAttributes(Assert.Single(syntax.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink).Attributes, "lnk", new[] { "primary" }, ("title", "Site"));
            AssertAttributes(Assert.Single(syntax.Children, node => node.Kind == MarkdownSyntaxKind.InlineImage).Attributes, "img", new[] { "wide" }, ("title", "Image"));
            AssertAttributes(Assert.Single(syntax.Children, node => node.Kind == MarkdownSyntaxKind.InlineEmphasis).Attributes, "em", new[] { "marked" });
            AssertAttributes(Assert.Single(syntax.Children, node => node.Kind == MarkdownSyntaxKind.InlineStrong).Attributes, "strong", new[] { "marked" });
            AssertAttributes(Assert.Single(syntax.Children, node => node.Kind == MarkdownSyntaxKind.InlineCodeSpan).Attributes, "code", new[] { "token" });

            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            Assert.Equal("<p><a href=\"https://example.com\" id=\"lnk\" class=\"primary\" title=\"Site\">site</a> <img src=\"image.png\" alt=\"alt\" id=\"img\" class=\"wide\" title=\"Image\" /> <em id=\"em\" class=\"marked\">emphasis</em> <strong id=\"strong\" class=\"marked\">strong</strong> <code id=\"code\" class=\"token\">code</code></p>", html);

            var written = result.Document.ToMarkdown().TrimEnd().Replace("\r\n", "\n");
            Assert.Equal(md, written);

            var reparsed = MarkdownReader.Parse(written, options);
            var reparsedParagraph = Assert.IsType<ParagraphBlock>(reparsed.Blocks[0]);
            Assert.Equal("lnk", Assert.Single(reparsedParagraph.Inlines.Nodes.OfType<LinkInline>()).Attributes.ElementId);
            Assert.Equal("img", Assert.Single(reparsedParagraph.Inlines.Nodes.OfType<ImageInline>()).Attributes.ElementId);
            Assert.Equal("em", Assert.Single(reparsedParagraph.Inlines.Nodes.OfType<ItalicSequenceInline>()).Attributes.ElementId);
            Assert.Equal("strong", Assert.Single(reparsedParagraph.Inlines.Nodes.OfType<BoldSequenceInline>()).Attributes.ElementId);
            Assert.Equal("code", Assert.Single(reparsedParagraph.Inlines.Nodes.OfType<CodeSpanInline>()).Attributes.ElementId);
        }

        [Fact]
        public void GenericAttributes_SpacedAttributeBlock_AfterInlineElement_TargetsParagraph() {
            const string md = "[site](https://example.com) {#lead .intro}";
            var options = new MarkdownReaderOptions {
                GenericAttributes = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(md, options);
            var paragraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[0]);
            var link = Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>());

            Assert.True(link.Attributes.IsEmpty);
            AssertAttributes(paragraph.Attributes, "lead", new[] { "intro" });
            Assert.Equal("[site](https://example.com)", paragraph.Inlines.RenderMarkdown());
        }

        [Fact]
        public void GenericAttributes_Are_Not_Parsed_WhenOptionDisabled() {
            const string md = "# Quarterly Revenue {#quarterly-overview .wide}";

            var result = MarkdownReader.ParseWithSyntaxTree(md);
            var heading = Assert.IsType<HeadingBlock>(result.Document.Blocks[0]);

            Assert.Equal("Quarterly Revenue {#quarterly-overview .wide}", heading.Text);
            Assert.True(heading.Attributes.IsEmpty);
            Assert.True(Assert.Single(result.SyntaxTree.Children).Attributes.IsEmpty);
        }

        [Fact]
        public void GenericAttributes_InlineAttributes_Are_Not_Parsed_WhenOptionDisabled() {
            const string md = "[site](https://example.com){#lnk .primary}";

            var result = MarkdownReader.ParseWithSyntaxTree(md);
            var paragraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[0]);
            var link = Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>());

            Assert.True(link.Attributes.IsEmpty);
            Assert.Contains(paragraph.Inlines.Nodes.OfType<TextRun>(), text => text.Text.Contains("{#lnk .primary}", StringComparison.Ordinal));
            Assert.Equal(md, paragraph.Inlines.RenderMarkdown());
        }

        private static void AssertAttributes(MarkdownAttributeSet attributes, string elementId, IReadOnlyList<string> classes, params (string Key, string Value)[] expectedAttributes) {
            Assert.Equal(elementId, attributes.ElementId);
            Assert.Equal(classes, attributes.Classes);
            foreach (var expected in expectedAttributes) {
                Assert.Equal(expected.Value, attributes.GetAttribute(expected.Key));
            }
        }

        [Fact]
        public void Fenced_Code_Block_Metadata_Can_Read_Typed_Boolean_And_Integer_Attributes() {
            const string md = """
```chart title="Quarterly Revenue" pinned compact=false maxItems=12 limit=7
{"series":[1,2,3]}
```
""";

            var doc = MarkdownReader.Parse(md);
            var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);

            Assert.True(code.FenceInfo.TryGetBooleanAttribute("pinned", out var pinned));
            Assert.True(pinned);
            Assert.True(code.FenceInfo.TryGetBooleanAttribute("compact", out var compact));
            Assert.False(compact);
            Assert.True(code.FenceInfo.TryGetInt32Attribute("maxItems", out var maxItems));
            Assert.Equal(12, maxItems);
            Assert.True(code.FenceInfo.TryGetInt32Attribute(out var aliasedLimit, "missing", "limit"));
            Assert.Equal(7, aliasedLimit);
            Assert.Equal("Quarterly Revenue", code.FenceInfo.GetAttribute("caption", "title"));
        }

        [Fact]
        public void Malformed_Fenced_Code_Block_Brace_Metadata_Does_Not_Partially_Apply_Structured_Attributes() {
            const string md = """
```chart {#quarterly-overview .wide title="Quarterly Revenue"
{"series":[1,2,3]}
```
""";

            var doc = MarkdownReader.Parse(md);
            var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);

            Assert.Equal("chart", code.Language);
            Assert.Equal("{#quarterly-overview .wide title=\"Quarterly Revenue\"", code.FenceInfo.AdditionalInfo);
            Assert.Null(code.FenceInfo.ElementId);
            Assert.Empty(code.FenceInfo.Classes);
            Assert.Null(code.FenceInfo.Title);
            Assert.False(code.FenceInfo.TryGetAttribute("title", out _));
            Assert.Equal(
                md.TrimEnd().Replace("\r\n", "\n"),
                ((IMarkdownBlock)code).RenderMarkdown().Replace("\r\n", "\n"));
        }

        [Fact]
        public void Fenced_Block_Extension_Matches_Primary_Language_When_InfoString_Has_Metadata() {
            var options = new MarkdownReaderOptions();
            options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
                "Chart semantic",
                new[] { "chart" },
                context => new SemanticFencedBlock(MarkdownSemanticKinds.Chart, context.InfoString, context.Content, context.Caption)));

            const string md = """
```chart title="Quarterly Revenue"
{"series":[1,2,3]}
```
""";

            var doc = MarkdownReader.Parse(md, options);
            var block = Assert.IsType<SemanticFencedBlock>(doc.Blocks[0]);

            Assert.Equal("chart", block.Language);
            Assert.Equal("chart title=\"Quarterly Revenue\"", block.InfoString);
            Assert.Equal("title=\"Quarterly Revenue\"", block.FenceInfo.AdditionalInfo);
            Assert.Equal("Quarterly Revenue", block.FenceInfo.Title);
            Assert.Equal(
                md.TrimEnd().Replace("\r\n", "\n"),
                ((IMarkdownBlock)block).RenderMarkdown().Replace("\r\n", "\n"));
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
            Assert.Equal(2, defList.InlineItems.Count);
            Assert.Equal("*Term*", defList.InlineItems[0].Term.RenderMarkdown());
            Assert.Equal("[Link](https://example.com)", defList.InlineItems[1].Term.RenderMarkdown());
            var html = ((IMarkdownBlock)defList).RenderHtml();
            Assert.Contains("<em>Term</em>", html);
            Assert.Contains("href=\"https://example.com\"", html);
        }

        [Fact]
        public void Paragraph_Exposes_Typed_Inline_Nodes_Alongside_Legacy_Items_View() {
            const string md = "[link](https://example.com)";

            var doc = MarkdownReader.Parse(md);
            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);

            var linkNode = Assert.Single(paragraph.Inlines.Nodes);
            var link = Assert.IsType<LinkInline>(linkNode);
            Assert.Equal("https://example.com", link.Url);

            var legacyItem = Assert.Single(paragraph.Inlines.Items);
            Assert.Same(linkNode, legacyItem);
        }

        [Fact]
        public void Definition_List_RenderHtml_Falls_Back_To_Current_StringItems_After_Mutation() {
            const string md = "Term: Value";

            var doc = MarkdownReader.Parse(md);
            var defList = Assert.IsType<DefinitionListBlock>(doc.Blocks[0]);

            defList.Items[0] = ("**Changed**", "[fresh](https://example.com)");

            var html = ((IMarkdownBlock)defList).RenderHtml();

            Assert.Contains("<dt><strong>Changed</strong></dt>", html, StringComparison.Ordinal);
            Assert.Contains("<dd><a href=\"https://example.com\">fresh</a></dd>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Definition_List_InlineItems_Follow_Current_String_Content_After_Mutation() {
            const string md = "Term: Value";

            var doc = MarkdownReader.Parse(md);
            var defList = Assert.IsType<DefinitionListBlock>(doc.Blocks[0]);

            defList.Items[0] = ("**Changed**", "[fresh](https://example.com)");

            Assert.Single(defList.InlineItems);
            Assert.Equal("**Changed**", defList.InlineItems[0].Term.RenderMarkdown());
            Assert.Equal("[fresh](https://example.com)", defList.InlineItems[0].Definition.RenderMarkdown());
        }

        [Fact]
        public void Definition_List_Does_Not_Consume_Literal_Url_Paragraphs() {
            string md = "Visit https://example.com/path_(x): now";
            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.DoesNotContain("<dl>", html, StringComparison.Ordinal);
            Assert.Contains("<p>Visit https://example.com/path_(x): now</p>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Definition_List_Can_Parse_Multi_Paragraph_Definition_Bodies() {
            string md = """
Term: First paragraph
  continued

  Second paragraph

Other: Value
""";

            var doc = MarkdownReader.Parse(md);

            var defList = Assert.IsType<DefinitionListBlock>(doc.Blocks[0]);
            Assert.Equal(2, defList.Entries.Count);
            Assert.Equal("Term", defList.Entries[0].Term.RenderMarkdown());
            Assert.Collection(defList.Entries[0].DefinitionBlocks,
                block => Assert.Equal("First paragraph\ncontinued", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.Equal("Second paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
            Assert.Equal("Other", defList.Entries[1].Term.RenderMarkdown());
            Assert.Equal("Value", Assert.IsType<ParagraphBlock>(Assert.Single(defList.Entries[1].DefinitionBlocks)).Inlines.RenderMarkdown());
        }

        [Fact]
        public void Definition_List_RenderMarkdown_Preserves_Multi_Block_Definition_Bodies_For_Reparse() {
            string md = """
Term: Intro

  - first
  - second

Other: Value
""";

            var doc = MarkdownReader.Parse(md);
            var roundTrip = MarkdownReader.Parse(doc.ToMarkdown());

            var defList = Assert.IsType<DefinitionListBlock>(Assert.Single(roundTrip.Blocks));
            Assert.Equal(2, defList.Entries.Count);
            Assert.Collection(defList.Entries[0].DefinitionBlocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<UnorderedListBlock>(block));
            Assert.Equal("Value", Assert.IsType<ParagraphBlock>(Assert.Single(defList.Entries[1].DefinitionBlocks)).Inlines.RenderMarkdown());
        }

        [Fact]
        public void Definition_List_Can_Parse_Nested_Blockquote_And_List_In_Definition_Body() {
            string md = """
Term: Intro

  > quoted line
  >
  > more quote

  - first
  - second
""";

            var doc = MarkdownReader.Parse(md);

            var definitions = Assert.IsType<DefinitionListBlock>(Assert.Single(doc.Blocks));
            var entry = Assert.Single(definitions.Entries);

            Assert.Collection(entry.DefinitionBlocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<QuoteBlock>(block),
                block => Assert.IsType<UnorderedListBlock>(block));
        }

        [Fact]
        public void Unordered_List_Item_With_Colon_Is_Not_Parsed_As_Definition_List() {
            string md = """
- **AD1**: starkes Muster (`7034/7023`).
- **AD2**: eher Secure-Channel (`3210/1129`).
""";

            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
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

        [Theory]
        [InlineData("<a href=\"foo  \nbar\">\n", "<p><a href=\"foo  \nbar\"></p>\n")]
        [InlineData("<a href=\"foo\\\nbar\">\n", "<p><a href=\"foo\\\nbar\"></p>\n")]
        public void CommonMark_Hard_Break_Markers_Inside_Raw_Inline_Html_Stay_Literal(string markdown, string expectedHtml) {
            var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
            var html = doc.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            Assert.Equal(CommonMarkHtmlComparison.Normalize(expectedHtml), CommonMarkHtmlComparison.Normalize(html));
        }

        [Fact]
        public void CommonMark_Trailing_Backslash_At_End_Of_Block_Stays_Literal() {
            const string markdown = "foo\\\n";
            const string expectedHtml = "<p>foo\\</p>\n";

            var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
            var html = doc.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            Assert.Equal(CommonMarkHtmlComparison.Normalize(expectedHtml), CommonMarkHtmlComparison.Normalize(html));
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
        public void Multiline_CodeSpan_Preserves_Source_Line_Ending_Content() {
            string md = string.Join("\n", new[] { "``", "foo ", "``", string.Empty });

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var code = Assert.IsType<CodeSpanInline>(Assert.Single(paragraph.Inlines.Nodes));

            Assert.Equal("foo ", code.Text);
            Assert.Contains("<code>foo </code>", doc.ToHtmlFragment());
        }

        [Fact]
        public void CodeSpan_Treats_Source_Backslash_Line_End_As_Literal_Content() {
            const string md = """
`code\
span`
""";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var code = Assert.IsType<CodeSpanInline>(Assert.Single(paragraph.Inlines.Nodes));

            Assert.Equal("code\\ span", code.Text);
            Assert.Contains("<code>code\\ span</code>", doc.ToHtmlFragment());
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
        public void Link_Label_With_Nested_Link_Deactivates_Outer_Link_Opener() {
            const string md = "[foo [bar](/uri)](/outer)";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var link = Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>());

            Assert.Equal("bar", link.Text);
            Assert.Equal("/uri", link.Url);
            Assert.Contains("[foo <a href=\"/uri\">bar</a>](/outer)", doc.ToHtmlFragment());
        }

        [Fact]
        public void Reference_Link_Label_Allows_Inline_Image_Content() {
            const string md = """
[![moon](moon.jpg)][ref]

[ref]: /uri
""";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var link = Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>());
            var image = Assert.Single(link.LabelInlines!.Nodes.OfType<ImageInline>());

            Assert.Equal("/uri", link.Url);
            Assert.Equal("moon.jpg", image.Src);
            Assert.Equal("moon", image.PlainAlt);
            Assert.Contains("<a href=\"/uri\"><img src=\"moon.jpg\" alt=\"moon\" /></a>", doc.ToHtmlFragment());
        }

        [Fact]
        public void Unresolved_Shortcut_Reference_Label_Reprocesses_Backslash_Escapes_As_Text() {
            const string md = """
            [bar][foo\!]

            [foo!]: /url
            """;

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.Single(doc.Blocks.OfType<ParagraphBlock>());

            Assert.DoesNotContain(paragraph.Inlines.Nodes, node => node is LinkInline);
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Equal("<p>[bar][foo!]</p>", html);
        }

        [Fact]
        public void Link_Label_Scanner_Ignores_Closing_Brackets_Inside_Code_Spans() {
            const string md = "[foo`](/uri)`";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var code = Assert.Single(paragraph.Inlines.Nodes.OfType<CodeSpanInline>());

            Assert.DoesNotContain(paragraph.Inlines.Nodes, node => node is LinkInline);
            Assert.Equal("](/uri)", code.Text);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Equal("<p>[foo<code>](/uri)</code></p>", html);
        }

        [Fact]
        public void Link_Label_Scanner_Ignores_Closing_Brackets_Inside_Angle_Autolinks() {
            const string md = "[foo<https://example.com/?search=](uri)>";

            var doc = MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var autolink = Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>());

            Assert.Equal("https://example.com/?search=](uri)", autolink.Text);
            Assert.Equal("https://example.com/?search=](uri)", autolink.Url);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("[foo<a href=\"https://example.com/?search=%5D(uri)\">", html);
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

        [Fact]
        public void Footnotes_Can_Parse_Nested_List_Blocks_And_RoundTrip() {
            string md = """
Lead[^a]

[^a]: Intro

  - first
  - second
""";

            var doc = MarkdownReader.Parse(md);
            var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(doc.Blocks, block => block is FootnoteDefinitionBlock));

            Assert.Collection(footnote.Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<UnorderedListBlock>(block));

            var roundTrip = MarkdownReader.Parse(doc.ToMarkdown());
            var reparsedFootnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(roundTrip.Blocks, block => block is FootnoteDefinitionBlock));
            Assert.Collection(reparsedFootnote.Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<UnorderedListBlock>(block));
        }
    }
}

