using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Profile_Tests {
    [Fact]
    public void Portable_Profile_Disables_OfficeImo_Only_Block_Extensions() {
        var options = MarkdownReaderOptions.CreatePortableProfile();

        Assert.False(options.Callouts);
        Assert.False(options.TaskLists);
        Assert.False(options.TocPlaceholders);
        Assert.False(options.Footnotes);
        Assert.False(options.StandaloneImageBlocks);
        Assert.False(options.AutolinkUrls);
        Assert.False(options.AutolinkWwwUrls);
        Assert.False(options.AutolinkEmails);
        Assert.True(options.Tables);
        Assert.True(options.DefinitionLists);
        Assert.Empty(options.BlockParserExtensions);
        Assert.Empty(options.InlineParserExtensions);
    }

    [Fact]
    public void CommonMark_Profile_Disables_Extensions_Beyond_Core_Syntax() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();

        Assert.False(options.FrontMatter);
        Assert.False(options.Callouts);
        Assert.False(options.TaskLists);
        Assert.False(options.Tables);
        Assert.False(options.DefinitionLists);
        Assert.False(options.TocPlaceholders);
        Assert.False(options.Footnotes);
        Assert.False(options.StandaloneImageBlocks);
        Assert.False(options.AutolinkUrls);
        Assert.False(options.AutolinkWwwUrls);
        Assert.False(options.AutolinkEmails);
        Assert.True(options.HtmlBlocks);
        Assert.True(options.InlineHtml);
        Assert.Empty(options.BlockParserExtensions);
        Assert.Empty(options.InlineParserExtensions);
    }

    [Fact]
    public void Gfm_Profile_Enables_Gfm_Syntax_But_Not_OfficeImo_Extensions() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();

        Assert.False(options.FrontMatter);
        Assert.False(options.Callouts);
        Assert.True(options.TaskLists);
        Assert.True(options.Tables);
        Assert.False(options.DefinitionLists);
        Assert.False(options.TocPlaceholders);
        Assert.True(options.Footnotes);
        Assert.False(options.StandaloneImageBlocks);
        Assert.True(options.SingleTildeStrikethrough);
        Assert.True(options.AutolinkUrls);
        Assert.True(options.AutolinkWwwUrls);
        Assert.Equal("http://", options.AutolinkWwwScheme);
        Assert.True(options.AutolinkEmails);
        Assert.Single(options.BlockParserExtensions);
        Assert.Empty(options.InlineParserExtensions);
        Assert.Equal(MarkdownReaderBuiltInExtensions.FootnotesExtensionName, options.BlockParserExtensions[0].Name);
    }

    [Fact]
    public void CreateProfile_Returns_Expected_Config() {
        var office = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO);
        var commonMark = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.CommonMark);
        var gfm = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.GitHubFlavoredMarkdown);
        var portable = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.Portable);

        Assert.True(office.Callouts);
        Assert.Equal(3, office.BlockParserExtensions.Count);
        Assert.Empty(office.InlineParserExtensions);
        Assert.False(commonMark.Callouts);
        Assert.Empty(commonMark.BlockParserExtensions);
        Assert.Empty(commonMark.InlineParserExtensions);
        Assert.True(gfm.Tables);
        Assert.Single(gfm.BlockParserExtensions);
        Assert.Empty(gfm.InlineParserExtensions);
        Assert.False(portable.Footnotes);
        Assert.False(portable.StandaloneImageBlocks);
        Assert.Empty(portable.BlockParserExtensions);
        Assert.Empty(portable.InlineParserExtensions);
    }

    [Fact]
    public void CommonMark_Profile_Keeps_Toc_And_Footnote_Syntax_As_Literal_Text() {
        const string markdown = """
[TOC]

Lead[^1]

[^1]: Footnote text
""";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        Assert.Equal(3, doc.Blocks.Count);

        var tocParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        Assert.Equal("\\[TOC\\]", tocParagraph.Inlines.RenderMarkdown());

        var leadParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        Assert.Equal("Lead\\[^1\\]", leadParagraph.Inlines.RenderMarkdown());

        var footnoteParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[2]);
        Assert.Equal("\\[^1\\]: Footnote text", footnoteParagraph.Inlines.RenderMarkdown());

        Assert.DoesNotContain(doc.Blocks, block => block is FootnoteDefinitionBlock);
    }

    [Fact]
    public void CommonMark_Profile_Supports_Multiline_Reference_Link_Definitions() {
        const string markdown = """
[foo]:
  /docs/start
  "Docs title"

[foo]
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"/docs/start\" title=\"Docs title\">foo</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Normalizes_Reference_Link_Labels_Case_Insensitively() {
        const string markdown = """
[FOO]: /url

[Foo]
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"/url\">Foo</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Percent_Encodes_Non_Ascii_Link_Destinations() {
        const string markdown = """
[foo]: /φου

[foo]
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"/%CF%86%CE%BF%CF%85\">foo</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Preserves_Bare_Autolink_Targets_Without_Trailing_Slash() {
        const string markdown = "<http://foo.bar.baz>";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"http://foo.bar.baz\">http://foo.bar.baz</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Supports_Multiline_Reference_Link_Labels() {
        const string markdown = """
[Foo
  bar]: /url

[Baz][Foo bar]
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"/url\">Baz</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Uses_Unicode_Case_Folding_For_Reference_Link_Labels() {
        const string markdown = """
[ẞ]
[SS]: /url
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"/url\">ẞ</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Falls_Back_To_Shortcut_References_When_Inline_Link_Syntax_Is_Invalid() {
        const string markdown = """
[foo](not a link)

[foo]: /url1
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p><a href=\"/url1\">foo</a>(not a link)</p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Backtracks_Unmatched_Reference_Like_Syntax_To_Later_Full_References() {
        const string markdown = """
[foo][bar][baz]

[baz]: /url
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p>[foo]<a href=\"/url\">bar</a></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Renders_Image_Alt_As_Plain_String_Content() {
        const string markdown = """
Lead ![foo *bar*](train.jpg "train & tracks")
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p>Lead <img src=\"train.jpg\" alt=\"foo bar\" title=\"train &amp; tracks\" /></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Flattens_Nested_Links_And_Images_Inside_Image_Alt_Text() {
        const string markdown = """
Lead ![foo ![bar](/url)](/url2) and ![foo [bar](/url)](/url2)
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p>Lead <img src=\"/url2\" alt=\"foo bar\" /> and <img src=\"/url2\" alt=\"foo bar\" /></p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Leaves_Standalone_Image_Lines_As_Paragraphs() {
        const string markdown = "![foo](train.jpg \"title\")";

        var document = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var image = Assert.Single(paragraph.Inlines.Nodes, node => node is ImageInline);
        Assert.Equal("foo", Assert.IsType<ImageInline>(image).Alt);
    }

    [Fact]
    public void OfficeImo_Profile_Keeps_Standalone_Image_Lines_As_ImageBlocks() {
        const string markdown = "![foo](train.jpg \"title\")";

        var document = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateOfficeIMOProfile());

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("foo", image.Alt);
    }

    [Fact]
    public void CommonMark_Profile_Allows_List_Items_To_Start_With_Indented_Code_Blocks() {
        const string markdown = """
1.     indented code
   paragraph

       more code
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ol><li><pre><code>indented code\n</code></pre><p>paragraph</p><pre><code>more code\n</code></pre></li></ol>", html);
    }

    [Fact]
    public void CommonMark_Profile_Preserves_Extra_Code_Indent_When_List_Item_Starts_With_Indented_Code() {
        const string markdown = """
1.      indented code
   paragraph

       more code
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ol><li><pre><code> indented code\n</code></pre><p>paragraph</p><pre><code>more code\n</code></pre></li></ol>", html);
    }

    [Fact]
    public void CommonMark_Profile_Renders_BlankStart_And_Empty_List_Items() {
        const string markdown = """
-
  foo
-
  ```
  bar
  ```
-
      baz

- foo
-
- bar

1. foo
2.
3. bar

*
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal(
            "<ul><li><p>foo</p></li><li><pre><code>bar\n</code></pre></li><li><pre><code>baz\n</code></pre></li><li><p>foo</p></li><li></li><li><p>bar</p></li></ul><ol><li>foo</li><li></li><li>bar</li></ol><ul><li></li></ul>",
            html);
    }

    [Fact]
    public void CommonMark_Profile_Does_Not_Let_Empty_List_Items_Interrupt_Paragraphs() {
        const string markdown = """
foo
*

foo
1.
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p>foo *</p><p>foo 1.</p>", html);
    }

    [Fact]
    public void CommonMark_Profile_Treats_Reference_Definition_List_Items_As_Loose_Without_Rendering_The_Definition() {
        const string markdown = """
- a
- b

  [ref]: /url
- d
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li><p>a</p></li><li><p>b</p></li><li><p>d</p></li></ul>", html);
    }

    [Fact]
    public void CommonMark_Profile_Keeps_Fenced_Code_Blank_Lines_Inside_A_Tight_List_Item() {
        const string markdown = """
- a
- ```
  b

  ```
- c
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li>a</li><li><pre><code>b\n\n</code></pre></li><li>c</li></ul>", html);
    }

    [Fact]
    public void CommonMark_Profile_Marks_Block_Leading_List_Items_As_Loose_When_A_Blank_Line_Separates_Block_Children() {
        const string markdown = """
1. ```
   foo
   ```

   bar
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ol><li><pre><code>foo\n</code></pre><p>bar</p></li></ol>", html);
    }

    [Fact]
    public void CommonMark_Profile_Requires_Full_Continuation_Indent_For_Nested_Lists_Under_Wide_Ordered_Markers() {
        const string markdown = """
10) foo
   - bar
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ol start=\"10\"><li>foo</li></ol><ul><li>bar</li></ul>", html);
    }

    [Fact]
    public void CommonMark_Profile_Does_Not_Nest_Shallowly_Indented_Sibling_List_Items() {
        const string markdown = """
- foo
 - bar
  - baz
   - boo
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li>foo</li><li>bar</li><li>baz</li><li>boo</li></ul>", html);
    }

    [Fact]
    public void CommonMark_Profile_Allows_Nested_Lists_As_The_First_Block_Inside_List_Items() {
        const string markdown = """
- - foo

1. - 2. foo
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li><ul><li>foo</li></ul></li></ul><ol><li><ul><li><ol start=\"2\"><li>foo</li></ol></li></ul></li></ol>", html);
    }

    [Fact]
    public void CommonMark_Profile_Allows_Headings_As_The_First_Block_Inside_List_Items() {
        const string markdown = """
- # Foo
- Bar
  ---
  baz
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li><h1 id=\"foo\">Foo</h1></li><li><h2 id=\"bar\">Bar</h2>baz</li></ul>", html);
    }

    [Fact]
    public void CommonMark_Profile_Starts_A_New_Bullet_List_When_The_Marker_Changes() {
        const string markdown = """
- foo
- bar
+ baz
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li>foo</li><li>bar</li></ul><ul><li>baz</li></ul>", html);
    }

    [Fact]
    public void CommonMark_Profile_Starts_A_New_Ordered_List_When_The_Delimiter_Changes() {
        const string markdown = """
1. foo
2. bar
3) baz
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ol><li>foo</li><li>bar</li></ol><ol start=\"3\"><li>baz</li></ol>", html);
    }

    [Fact]
    public void CommonMark_Profile_Keeps_Blank_Line_Separated_Items_In_A_Single_Loose_List() {
        const string markdown = """
- foo

- bar

- baz
""";

        var html = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<ul><li><p>foo</p></li><li><p>bar</p></li><li><p>baz</p></li></ul>", html);
    }

    [Fact]
    public void Gfm_Profile_Parses_Tables_TaskLists_And_Footnotes_But_Not_Toc_Placeholders() {
        const string markdown = """
[TOC]

- [x] Done

| Col |
| --- |
| Value |

Lead[^1]

[^1]: Footnote text
""";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

        Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        Assert.Equal("\\[TOC\\]", Assert.IsType<ParagraphBlock>(doc.Blocks[0]).Inlines.RenderMarkdown());

        var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[1]);
        Assert.True(list.Items[0].IsTask);
        Assert.True(list.Items[0].Checked);

        Assert.IsType<TableBlock>(doc.Blocks[2]);
        Assert.IsType<ParagraphBlock>(doc.Blocks[3]);
        Assert.IsType<FootnoteDefinitionBlock>(doc.Blocks[4]);
    }

    [Fact]
    public void CommonMark_Profile_Can_Opt_Into_Callout_Extension_Explicitly() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.Callouts = true;
        MarkdownReaderBuiltInExtensions.AddCallouts(options);

        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""", options);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("note", callout.Kind);
        Assert.Equal("Example", callout.TitleInlines.RenderMarkdown());
    }

    [Fact]
    public void Gfm_Profile_Enables_SingleTilde_Strikethrough_While_CommonMark_Keeps_It_Literal() {
        const string markdown = "A proper ~strikethrough~.";

        var gfmHtml = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
        var commonMarkHtml = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile())
            .ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });

        Assert.Equal("<p>A proper <del>strikethrough</del>.</p>", gfmHtml);
        Assert.Equal("<p>A proper ~strikethrough~.</p>", commonMarkHtml);
    }

    [Fact]
    public void Gfm_Html_Option_Renders_Task_Lists_Without_OfficeImo_Task_Classes() {
        var doc = MarkdownReader.Parse("- [ ] foo\n- [x] bar\n", MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

        var html = doc.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            GitHubTaskListHtml = true
        });

        Assert.Equal(
            "<ul><li><input type=\"checkbox\" disabled=\"\" /> foo</li><li><input type=\"checkbox\" checked=\"\" disabled=\"\" /> bar</li></ul>",
            html);
        Assert.DoesNotContain("task-list-item", html, StringComparison.Ordinal);
        Assert.DoesNotContain("contains-task-list", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Html_Option_Renders_Github_Footnotes_And_Leaves_Missing_References_Literal() {
        const string markdown = """
Alpha[^1].

Missing[^nope].

[^1]: Note
""";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

        var html = doc.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            GitHubFootnoteHtml = true
        });

        Assert.Contains(
            "<sup class=\"footnote-ref\"><a href=\"#fn-1\" id=\"fnref-1\" data-footnote-ref>1</a></sup>",
            html,
            StringComparison.Ordinal);
        Assert.Contains("Missing[^nope].", html, StringComparison.Ordinal);
        Assert.Contains(
            "<section class=\"footnotes\" data-footnotes><ol><li id=\"fn-1\"><p>Note <a href=\"#fnref-1\" class=\"footnote-backref\" data-footnote-backref data-footnote-backref-idx=\"1\" aria-label=\"Back to reference 1\">↩</a></p></li></ol></section>",
            html,
            StringComparison.Ordinal);
        Assert.DoesNotContain("fn-nope", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Profile_Treats_Exclamation_As_Punctuation_Before_Footnote_Reference() {
        var doc = MarkdownReader.Parse(
            "This is some text![^1].\n\n[^1]: Note",
            MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

        var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);

        Assert.Collection(
            paragraph.Inlines.Nodes,
            node => Assert.Equal("This is some text", Assert.IsType<TextRun>(node).Text),
            node => Assert.Equal("!", Assert.IsType<TextRun>(node).Text),
            node => Assert.Equal("1", Assert.IsType<FootnoteRefInline>(node).Label),
            node => Assert.Equal(".", Assert.IsType<TextRun>(node).Text));
    }

    [Fact]
    public void Gfm_Profile_Lets_Footnote_Definitions_Interrupt_Paragraphs() {
        const string markdown = """
[^1]: Footnote body
Lead paragraph
[^footnote]:
    > Blockquotes can be in a footnote.

        as well as code blocks

    or, naturally, simple paragraphs.
""";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

        Assert.Collection(
            doc.Blocks,
            block => Assert.Equal("1", Assert.IsType<FootnoteDefinitionBlock>(block).Label),
            block => Assert.IsType<ParagraphBlock>(block),
            block => {
                var footnote = Assert.IsType<FootnoteDefinitionBlock>(block);
                Assert.Equal("footnote", footnote.Label);
                Assert.Contains(footnote.Blocks, nested => nested is QuoteBlock);
                Assert.Contains(footnote.Blocks, nested => nested is CodeBlock);
                Assert.Contains(footnote.Blocks, nested => nested is ParagraphBlock);
            });
    }
}
