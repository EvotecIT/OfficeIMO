using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Refs_Footnotes_Tests {
        [Fact]
        public void Reference_Links_Are_Resolved() {
            var md = string.Join("\n", new[] {
                "See [Docs][docs] and [Site][site].",
                "",
                "[docs]: https://evotec.xyz \"Docs\"",
                "[site]: <https://example.com> \"Site\""
            });
            var doc = MarkdownReader.Parse(md);
            var outMd = doc.ToMarkdown();
            // Either inline links or preserved, accept either; primarily ensure resolution in HTML
            var html = doc.ToHtml();
            Assert.Contains("https://evotec.xyz", html);
            Assert.Contains("https://example.com", html);
            Assert.DoesNotContain("[docs]:", outMd); // definitions consumed
        }

        [Fact]
        public void Reference_Links_With_Nested_Label_Text_Are_Resolved() {
            var md = string.Join("\n", new[] {
                "See [Docs [API]][docs].",
                "",
                "[docs]: https://evotec.xyz"
            });

            var html = MarkdownReader.Parse(md).ToHtml();

            Assert.Contains("href=\"https://evotec.xyz\"", html);
            Assert.Contains(">Docs [API]<", html);
        }

        [Fact]
        public void Supported_Html_Wrappers_In_Link_Labels_Do_Not_Reenable_Nested_Links() {
            const string md = "[outer <u>[inner](https://inner.example)</u>](https://outer.example)";

            var doc = MarkdownReader.Parse(md);
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var link = Assert.IsType<LinkInline>(Assert.Single(paragraph.Inlines.Nodes));
            Assert.NotNull(link.LabelInlines);

            var wrapper = Assert.IsType<HtmlTagSequenceInline>(Assert.Single(link.LabelInlines!.Nodes, node => node is HtmlTagSequenceInline));
            Assert.DoesNotContain(wrapper.Inlines.Nodes, node => node is LinkInline);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

            Assert.Contains("href=\"https://outer.example\"", html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"https://inner.example\"", html, StringComparison.Ordinal);
            Assert.Contains("[inner](https://inner.example)", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Supported_Html_Wrappers_Decode_Entities_Before_Parsing_Inlines() {
            const string md = "Value <u>&amp;</u>";

            var doc = MarkdownReader.Parse(md);
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var wrapper = Assert.IsType<HtmlTagSequenceInline>(Assert.Single(paragraph.Inlines.Nodes, node => node is HtmlTagSequenceInline));
            var text = Assert.IsType<DecodedHtmlEntityTextRun>(Assert.Single(wrapper.Inlines.Nodes));

            Assert.Equal("&", text.Text);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("Value <u>&amp;</u>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("&amp;amp;", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Supported_Html_Wrappers_Keep_Decoded_Tag_Text_Literal() {
            const string md = "Value <u>&lt;u&gt;x&lt;/u&gt;</u>";

            var doc = MarkdownReader.Parse(md);
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var wrapper = Assert.IsType<HtmlTagSequenceInline>(Assert.Single(paragraph.Inlines.Nodes, node => node is HtmlTagSequenceInline));

            Assert.Contains(wrapper.Inlines.Nodes, node => node is DecodedHtmlEntityTextRun);
            Assert.All(wrapper.Inlines.Nodes, node => Assert.True(node is TextRun or DecodedHtmlEntityTextRun, node.GetType().FullName));
            Assert.Equal("<u>x</u>", InlinePlainText.Extract(wrapper.Inlines));

            var markdown = doc.ToMarkdown();
            Assert.Contains("Value <u>&lt;u&gt;x&lt;/u&gt;</u>", markdown, StringComparison.Ordinal);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("Value <u>&lt;u&gt;x&lt;/u&gt;</u>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Supported_Html_Wrappers_Keep_Decoded_Markdown_Delimiters_Literal() {
            const string md = "Value <u>&#96;code&#96; &#126;&#126;strike&#126;&#126; &#61;&#61;mark&#61;&#61;</u>";

            var doc = MarkdownReader.Parse(md);
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var wrapper = Assert.IsType<HtmlTagSequenceInline>(Assert.Single(paragraph.Inlines.Nodes, node => node is HtmlTagSequenceInline));

            Assert.Contains(wrapper.Inlines.Nodes, node => node is DecodedHtmlEntityTextRun);
            Assert.All(wrapper.Inlines.Nodes, node => Assert.True(node is TextRun or DecodedHtmlEntityTextRun, node.GetType().FullName));
            Assert.Equal("`code` ~~strike~~ ==mark==", InlinePlainText.Extract(wrapper.Inlines));

            var markdown = doc.ToMarkdown();
            Assert.Contains(@"Value <u>\`code\` \~\~strike\~\~ \=\=mark\=\=</u>", markdown, StringComparison.Ordinal);

            var reparsed = MarkdownReader.Parse(markdown);
            var html = reparsed.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("Value <u>`code` ~~strike~~ ==mark==</u>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<code>code</code>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<del>strike</del>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<mark>mark</mark>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Footnote_RenderMarkdown_Indents_NonParagraph_First_Block_On_Next_Line() {
            var footnote = new FootnoteDefinitionBlock(
                "1",
                "```csharp\nConsole.WriteLine(1);\n```",
                new IMarkdownBlock[] { new CodeBlock("csharp", "Console.WriteLine(1);") },
                syntaxChildren: null);

            string markdown = ((IMarkdownBlock)footnote).RenderMarkdown();

            Assert.StartsWith("[^1]:\n  ```csharp", markdown, StringComparison.Ordinal);
            Assert.Contains("\n  Console.WriteLine(1);\n", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void Footnote_Public_Structured_Constructor_Uses_ChildBlocks_As_Primary_Content() {
            var paragraph = new ParagraphBlock(MarkdownReader.ParseInlineText("Intro"));
            var list = new UnorderedListBlock();
            list.Items.Add(ListItem.Text("first"));

            var footnote = new FootnoteDefinitionBlock("audit", new IMarkdownBlock[] { paragraph, list });

            Assert.Equal(2, footnote.ChildBlocks.Count);
            Assert.Same(paragraph, footnote.ChildBlocks[0]);
            Assert.Same(list, footnote.ChildBlocks[1]);
            Assert.Equal(footnote.ChildBlocks, footnote.Blocks);
            Assert.Equal(footnote.ChildBlocks, ((IChildMarkdownBlockContainer)footnote).ChildBlocks);
            Assert.Equal("Intro\n\n- first", footnote.Text.Replace("\r\n", "\n"));
            Assert.Equal("[^audit]: Intro\n\n  - first", ((IMarkdownBlock)footnote).RenderMarkdown().Replace("\r\n", "\n"));
        }

        [Fact]
        public void Footnote_Public_Text_Constructor_Adapts_Text_To_ChildBlocks() {
            var footnote = new FootnoteDefinitionBlock("audit", "Intro *value*");

            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(footnote.ChildBlocks));

            Assert.Equal(footnote.ChildBlocks, footnote.Blocks);
            Assert.Equal(footnote.ChildBlocks, ((IChildMarkdownBlockContainer)footnote).ChildBlocks);
            Assert.Same(paragraph, Assert.Single(footnote.ParagraphBlocks));
            Assert.Same(paragraph.Inlines, Assert.Single(footnote.Paragraphs));
            Assert.Equal("Intro *value*", footnote.Text);
            Assert.Equal("Intro *value*", paragraph.Inlines.RenderMarkdown());
            Assert.Contains("<em>value</em>", ((IMarkdownBlock)footnote).RenderHtml(), StringComparison.Ordinal);
        }

        [Fact]
        public void Footnote_RenderMarkdown_Roundtrips_Structured_Body_With_Gfm_Profile() {
            var quote = new QuoteBlock();
            quote.Children.Add(new ParagraphBlock(MarkdownReader.ParseInlineText("Quoted *note*")));

            var footnote = new FootnoteDefinitionBlock(
                "shape",
                new IMarkdownBlock[] {
                    new ParagraphBlock(MarkdownReader.ParseInlineText("Intro *value*")),
                    quote,
                    new CodeBlock("text", "line 1\nline 2")
                });

            var doc = MarkdownDoc.Create()
                .Add(new ParagraphBlock(MarkdownReader.ParseInlineText("See [^shape].", MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile())))
                .Add(footnote);

            string markdown = doc.ToMarkdown().Replace("\r\n", "\n");

            Assert.Contains("See [^shape].", markdown, StringComparison.Ordinal);
            Assert.Contains("[^shape]: Intro *value*", markdown, StringComparison.Ordinal);
            Assert.Contains("\n\n  > Quoted *note*", markdown, StringComparison.Ordinal);
            Assert.Contains("\n\n  ```text\n  line 1\n  line 2\n  ```", markdown, StringComparison.Ordinal);

            var reparsed = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
            var reparsedFootnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(reparsed.Blocks, block => block is FootnoteDefinitionBlock));

            Assert.Collection(
                reparsedFootnote.Blocks,
                block => Assert.Equal("Intro *value*", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var reparsedQuote = Assert.IsType<QuoteBlock>(block);
                    var quotedParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(reparsedQuote.Children));
                    Assert.Equal("Quoted *note*", quotedParagraph.Inlines.RenderMarkdown());
                },
                block => {
                    var code = Assert.IsType<CodeBlock>(block);
                    Assert.Equal("text", code.InfoString);
                    Assert.Equal("line 1\nline 2", code.Content);
                });

            string html = reparsed.ToHtmlFragment(GfmHtmlComparison.CreatePlainHtmlOptions());

            Assert.Contains("<section class=\"footnotes\" data-footnotes>", html, StringComparison.Ordinal);
            Assert.Contains("id=\"fn-shape\"", html, StringComparison.Ordinal);
            Assert.Contains("<blockquote><p>Quoted <em>note</em></p></blockquote>", html, StringComparison.Ordinal);
            Assert.Contains("<pre><code class=\"language-text\">line 1\nline 2\n</code></pre>", html, StringComparison.Ordinal);

            int codeIndex = html.IndexOf("<pre><code class=\"language-text\">", StringComparison.Ordinal);
            int backrefIndex = html.IndexOf("href=\"#fnref-shape\"", StringComparison.Ordinal);
            Assert.True(codeIndex >= 0, html);
            Assert.True(backrefIndex > codeIndex, html);
        }

        [Fact]
        public void Reference_Link_Title_On_Next_Line_Is_Resolved() {
            var md = string.Join("\n", new[] {
                "See [Docs][docs].",
                "",
                "[docs]: https://evotec.xyz",
                "  \"Docs\""
            });

            var html = MarkdownReader.Parse(md).ToHtml();

            Assert.Contains("href=\"https://evotec.xyz\"", html);
            Assert.Contains("title=\"Docs\"", html);
            Assert.DoesNotContain("&quot;Docs&quot;", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Reference_Link_First_Definition_Wins() {
            var md = string.Join("\n", new[] {
                "See [Docs][docs].",
                "",
                "[docs]: https://first.example.com \"First\"",
                "[docs]: https://second.example.com \"Second\""
            });

            var html = MarkdownReader.Parse(md).ToHtml();

            Assert.Contains("href=\"https://first.example.com\"", html);
            Assert.Contains("title=\"First\"", html);
            Assert.DoesNotContain("https://second.example.com", html, StringComparison.Ordinal);
            Assert.DoesNotContain("title=\"Second\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Invalid_Nested_Reference_Definition_Line_Remains_Literal_Text() {
            const string md = "[x [y]]\n\n[x [y]]: https://example.com";

            var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

            Assert.Contains("<p>[x [y]]</p>", html, StringComparison.Ordinal);
            Assert.Contains("<p>[x [y]]: https://example.com</p>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("[]: https://example.com", "<p>[]: https://example.com</p>")]
        [InlineData("[ ]: https://example.com", "<p>[ ]: https://example.com</p>")]
        public void Invalid_Empty_Reference_Definition_Lines_Remain_Literal_Text(string markdown, string expectedHtml) {
            var html = MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

            Assert.Contains(expectedHtml, html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<dl>", html, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("[x]: https://example.com \"title\" extra", "<p>[x]: https://example.com &quot;title&quot; extra</p>")]
        [InlineData("[x]: <https://example.com/a b> \"title\" extra", "<p>[x]: &lt;https://example.com/a b&gt; &quot;title&quot; extra</p>")]
        public void Invalid_Reference_Definition_Title_Tails_Remain_Literal_Text(string markdown, string expectedHtml) {
            var html = MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

            Assert.Contains(expectedHtml, html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"https://example.com", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Footnote_Refs_And_Definitions_RoundTrip() {
            var md = string.Join("\n", new[] {
                "Hello[^1] world.",
                "",
                "[^1]: A note",
            });
            var doc = MarkdownReader.Parse(md);
            var outMd = doc.ToMarkdown();
            Assert.Contains("[^1]", outMd);
            Assert.Contains("[^1]: A note", outMd);
            var html = doc.ToHtml();
            Assert.Contains("id=\"fnref:1\"", html);
            Assert.Contains("id=\"fn:1\"", html);
        }

        [Fact]
        public void Standalone_Footnote_Block_Reuses_Parsed_Paragraphs_When_Available() {
            var md = string.Join("\n", new[] {
                "Lead[^1]",
                "",
                "[^1]: First *line*",
                "",
                "  Second [link](https://example.com)"
            });

            var doc = MarkdownReader.Parse(md);
            var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(doc.Blocks, b => b is FootnoteDefinitionBlock));

            var html = ((IMarkdownBlock)footnote).RenderHtml();

            Assert.Contains("<em>line</em>", html, StringComparison.Ordinal);
            Assert.Contains("href=\"https://example.com\"", html, StringComparison.Ordinal);
            Assert.Equal(2, footnote.Paragraphs.Count);
            Assert.Equal(2, footnote.ParagraphBlocks.Count);
            Assert.Same(footnote.Blocks[0], footnote.ParagraphBlocks[0]);
            Assert.Same(footnote.Blocks[1], footnote.ParagraphBlocks[1]);
            Assert.Same(footnote.ParagraphBlocks[0].Inlines, footnote.Paragraphs[0]);
            Assert.Same(footnote.ParagraphBlocks[1].Inlines, footnote.Paragraphs[1]);
            Assert.All(footnote.ParagraphBlocks, paragraph => Assert.False(string.IsNullOrWhiteSpace(paragraph.Inlines.RenderMarkdown())));
        }

        [Fact]
        public void Footnote_Text_Is_Derived_From_Blocks_When_BlockContent_Is_Available() {
            var paragraph = new ParagraphBlock(MarkdownReader.ParseInlineText("fresh value"));
            var footnote = new FootnoteDefinitionBlock("1", "stale value", new IMarkdownBlock[] { paragraph }, syntaxChildren: null);

            Assert.Equal("fresh value", footnote.Text);
            Assert.Same(paragraph, Assert.Single(footnote.Blocks));
            Assert.Same(paragraph, Assert.Single(footnote.ParagraphBlocks));
            Assert.Same(paragraph.Inlines, Assert.Single(footnote.Paragraphs));
        }

        [Fact]
        public void Footnote_Syntax_Rebuilds_When_Cached_Children_Do_Not_Match_Canonical_Blocks() {
            var staleParagraph = new ParagraphBlock(MarkdownReader.ParseInlineText("stale value"));
            var staleSyntax = ((ISyntaxMarkdownBlock)staleParagraph).BuildSyntaxNode(new MarkdownSourceSpan(1, 7, 1, 17));
            var freshParagraph = new ParagraphBlock(MarkdownReader.ParseInlineText("fresh value"));
            var footnote = new FootnoteDefinitionBlock(
                "1",
                "fallback value",
                new IMarkdownBlock[] { freshParagraph },
                new[] { staleSyntax });

            var syntax = ((ISyntaxMarkdownBlock)footnote).BuildSyntaxNode(new MarkdownSourceSpan(1, 1, 1, 18));
            var paragraphSyntax = Assert.Single(syntax.Children, child => child.Kind == MarkdownSyntaxKind.Paragraph);

            Assert.Same(freshParagraph, paragraphSyntax.AssociatedObject);
            Assert.Equal("fresh value", paragraphSyntax.Literal);
            Assert.DoesNotContain(syntax.Children, child => ReferenceEquals(child.AssociatedObject, staleParagraph));
        }
    }
}
