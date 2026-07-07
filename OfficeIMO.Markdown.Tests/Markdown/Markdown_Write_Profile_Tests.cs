using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Write_Profile_Tests {
    [Fact]
    public void Portable_Write_Profile_Degrades_Callouts_To_Quoted_Markdown() {
        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""");

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreatePortableProfile()).Replace("\r\n", "\n");

        Assert.DoesNotContain("[!NOTE]", markdown, StringComparison.Ordinal);
        Assert.Contains("> **Example**", markdown, StringComparison.Ordinal);
        Assert.Contains("> Body text", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Portable_Callout_Html_Fallback_Removes_OfficeImo_Callout_Chrome() {
        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""");
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };
        MarkdownBlockRenderBuiltInExtensions.AddPortableCalloutHtmlFallback(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<blockquote>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<strong>Example</strong>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Body text</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"callout", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Markdig_Alert_Html_Fallback_Renders_GitHub_Alert_Chrome() {
        var doc = MarkdownReader.Parse("""
> [!NOTE]
> Body text
""");
        var options = new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<div class=\"markdown-alert markdown-alert-note\">", html, StringComparison.Ordinal);
        Assert.Contains("class=\"markdown-alert-title\"", html, StringComparison.Ordinal);
        Assert.Contains(">Note</p>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Body text</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"callout", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Markdig_Alert_Html_Fallback_Leaves_Custom_Kinds_Untitled() {
        var doc = MarkdownReader.Parse("""
> [!CUSTOM]
> Body text
""");
        var options = new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<div class=\"markdown-alert markdown-alert-custom\">", html, StringComparison.Ordinal);
        Assert.Contains("<p>Body text</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("markdown-alert-title", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdig_Alert_Html_Fallback_Preserves_OfficeImo_Title_Inlines() {
        var doc = MarkdownReader.Parse("""
> [!NOTE] **Example**
> Body text
""");
        var options = new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<div class=\"markdown-alert markdown-alert-note\">", html, StringComparison.Ordinal);
        Assert.Contains("<p class=\"markdown-alert-title\"><svg", html, StringComparison.Ordinal);
        Assert.Contains("<strong>Example</strong>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Body text</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdig_Alert_Html_Fallback_Is_Registered_Once() {
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };

        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(options);

        Assert.Single(
            options.BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.MarkdigAlertHtmlName, StringComparison.Ordinal));
    }

    [Fact]
    public void Portable_Html_Fallbacks_Render_Toc_As_Plain_List() {
        var doc = MarkdownDoc.Create()
            .H2("Section")
            .H3("Child")
            .TocHere(options => {
                options.IncludeTitle = true;
                options.Title = "Contents";
                options.TitleLevel = 2;
                options.Layout = TocLayout.Panel;
            });
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };
        MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<h2>Contents</h2>", html, StringComparison.Ordinal);
        Assert.Contains("<ul>", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#section\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"md-toc", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<nav", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Portable_Html_Fallbacks_Render_Footnotes_Without_OfficeImo_Section_Chrome() {
        var doc = MarkdownReader.Parse("""
Lead[^1]

[^1]: Footnote text
""");
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };
        MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<section><hr />", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p id=\"fn:1\"><sup>1</sup> Footnote text", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"footnotes\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<ol>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Portable_Write_Profile_Omits_OfficeImo_Image_Size_Suffixes() {
        var doc = MarkdownDoc.Create()
            .Add(new ImageBlock("https://example.com/logo.png", "Logo", "Example", width: 256, height: 128));

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreatePortableProfile()).Replace("\r\n", "\n").Trim();

        Assert.Equal("![Logo](https://example.com/logo.png \"Example\")", markdown);
        Assert.DoesNotContain("{width=", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Tables_As_Raw_Html() {
        var table = new TableBlock();
        table.Headers.Add("Name");
        table.Headers.Add("Value");
        table.Rows.Add(new[] { "Area", "Markdown" });
        var doc = MarkdownDoc.Create().Add(table);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.StartsWith("<table>", markdown, StringComparison.Ordinal);
        Assert.Contains("<th>Name</th>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("| Name | Value |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Gfm_Only_Inlines_As_Raw_Html() {
        var paragraph = new ParagraphBlock(new InlineSequence()
            .Strike("old")
            .Text("and")
            .Highlight("new"));
        var doc = MarkdownDoc.Create().Add(paragraph);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Equal("<del>old</del> and <mark>new</mark>", markdown);
        Assert.DoesNotContain("~~old~~", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("==new==", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Task_Lists_As_Raw_Html() {
        var list = new UnorderedListBlock();
        list.Items.Add(ListItem.Task("Done", done: true));
        list.Items.Add(ListItem.Task("Open"));
        var doc = MarkdownDoc.Create().Add(list);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.StartsWith("<ul", markdown, StringComparison.Ordinal);
        Assert.Contains("type=\"checkbox\"", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- [x]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- [ ]", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Definition_Lists_As_Raw_Html() {
        var list = new DefinitionListBlock();
        list.AddEntry(new DefinitionListEntry(
            new InlineSequence().Text("Term"),
            new IMarkdownBlock[] { new ParagraphBlock(new InlineSequence().Text("Definition")) }));
        var doc = MarkdownDoc.Create().Add(list);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.StartsWith("<dl>", markdown, StringComparison.Ordinal);
        Assert.Contains("<dt>Term</dt>", markdown, StringComparison.Ordinal);
        Assert.Contains("<dd>Definition</dd>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(": Definition", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Footnotes_As_Raw_Html() {
        var doc = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence().Text("Lead").FootnoteRef("1")))
            .Add(new FootnoteDefinitionBlock("1", "Footnote text"));

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Contains("<sup id=\"fnref:1\"><a href=\"#fn:1\">1</a></sup>", markdown, StringComparison.Ordinal);
        Assert.Contains("<p id=\"fn:1\"><sup>1</sup> Footnote text", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[^1]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[^1]:", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Details_As_Raw_Html() {
        var details = new DetailsBlock(
            new SummaryBlock("More"),
            new IMarkdownBlock[] {
                new ParagraphBlock(new InlineSequence().Bold("Hidden"))
            },
            open: true);
        var doc = MarkdownDoc.Create().Add(details);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.StartsWith("<details open>", markdown, StringComparison.Ordinal);
        Assert.Contains("<summary>More</summary>", markdown, StringComparison.Ordinal);
        Assert.Contains("<strong>Hidden</strong>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("**Hidden**", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_CustomContainers_As_Raw_Html() {
        var doc = MarkdownDoc.Create().Add(new CustomContainerBlock(
            "note",
            new IMarkdownBlock[] {
                new ParagraphBlock(new InlineSequence().Text("Body"))
            }));

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Equal("<div class=\"note\"><p>Body</p></div>", markdown);
        Assert.DoesNotContain(":::", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Escapes_DirectParagraphLineStarts() {
        var doc = MarkdownDoc.Create().Add(new ParagraphBlock(new InlineSequence()
            .Text("# not heading")
            .SoftBreak()
            .Text("- not list")
            .SoftBreak()
            .Text("1. not ordered")));

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Equal("\\# not heading\n\\- not list\n1\\. not ordered", markdown);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_Attributed_Blocks_As_Raw_Html() {
        var readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        readerOptions.GenericAttributes = true;
        var doc = MarkdownReader.Parse("""
# Title {#intro .lead}

{#para .note}
Text
""", readerOptions);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Contains("<h1 id=\"intro\" class=\"lead\">Title</h1>", markdown, StringComparison.Ordinal);
        Assert.Contains("<p id=\"para\" class=\"note\">Text</p>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("{#intro", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("{#para", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_NonCommonMark_Inlines_As_Raw_Html() {
        var scalarDoc = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence()
                .Inserted("added")
                .Superscript("up")
                .Subscript("down")));

        var scalarMarkdown = scalarDoc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Equal("<ins>added</ins> <sup>up</sup> <sub>down</sub>", scalarMarkdown);

        var readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        readerOptions.Inserted = true;
        readerOptions.Superscript = true;
        readerOptions.Subscript = true;
        var sequenceDoc = MarkdownReader.Parse("++added **bold**++ ^up **bold**^ ~down **bold**~", readerOptions);

        var sequenceMarkdown = sequenceDoc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Contains("<ins>added <strong>bold</strong></ins>", sequenceMarkdown, StringComparison.Ordinal);
        Assert.Contains("<sup>up <strong>bold</strong></sup>", sequenceMarkdown, StringComparison.Ordinal);
        Assert.Contains("<sub>down <strong>bold</strong></sub>", sequenceMarkdown, StringComparison.Ordinal);
        Assert.DoesNotContain("++added", sequenceMarkdown, StringComparison.Ordinal);
        Assert.DoesNotContain("^up", sequenceMarkdown, StringComparison.Ordinal);
        Assert.DoesNotContain("~down", sequenceMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonMark_Write_Profile_Renders_InlineFallbackAttributes_As_HtmlAttributes() {
        var readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        readerOptions.GenericAttributes = true;
        var doc = MarkdownReader.Parse("~~old~~{#gone .muted}", readerOptions);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateCommonMarkProfile()).Replace("\r\n", "\n").Trim();

        Assert.Equal("<del id=\"gone\" class=\"muted\">old</del>", markdown);
        Assert.DoesNotContain("{#gone", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void GitHubFlavoredMarkdown_Write_Profile_Keeps_Pipe_Tables_And_Portable_Images() {
        var table = new TableBlock();
        table.Headers.Add("Name");
        table.Headers.Add("Value");
        table.Rows.Add(new[] { "Area", "Markdown" });
        var doc = MarkdownDoc.Create()
            .Add(new ImageBlock("https://example.com/logo.png", "Logo", width: 256, height: 128))
            .Add(table);

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateGitHubFlavoredMarkdownProfile()).Replace("\r\n", "\n").Trim();

        Assert.Contains("![Logo](https://example.com/logo.png)", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("{width=", markdown, StringComparison.Ordinal);
        Assert.Contains("| Name | Value |", markdown, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownWriteOptions_CreateProfile_Maps_Named_Output_Profiles() {
        Assert.Empty(MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.OfficeIMO).BlockRenderExtensions);
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkTableMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkTaskListMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkDefinitionListMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkFootnoteDefinitionMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkDetailsMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkCustomContainerMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkParagraphLineStartMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.CommonMarkAttributedBlockMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).InlineRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownInlineRenderBuiltInExtensions.CommonMarkStrikethroughMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).InlineRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownInlineRenderBuiltInExtensions.CommonMarkFootnoteReferenceMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).InlineRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownInlineRenderBuiltInExtensions.CommonMarkInsertedMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).InlineRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownInlineRenderBuiltInExtensions.CommonMarkSuperscriptMarkdownName, StringComparison.Ordinal));
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.CommonMark).InlineRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownInlineRenderBuiltInExtensions.CommonMarkSubscriptMarkdownName, StringComparison.Ordinal));
        Assert.Equal(MarkdownImageRenderingMode.PortableMarkdown, MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.GitHubFlavoredMarkdown).ImageRenderingMode);
        Assert.Contains(
            MarkdownWriteOptions.CreateProfile(MarkdownOutputProfile.Portable).BlockRenderExtensions,
            extension => string.Equals(extension.Name, MarkdownBlockRenderBuiltInExtensions.PortableCalloutMarkdownName, StringComparison.Ordinal));
    }

    [Fact]
    public void Html_Image_Profile_Renders_Linked_Image_Blocks_As_Raw_Html() {
        var doc = MarkdownDoc.Create()
            .Add(new ImageBlock(
                "https://example.com/logo.png",
                "Logo",
                "Example",
                width: 256,
                height: 128,
                linkUrl: "https://example.com/docs",
                linkTitle: "Documentation",
                linkTarget: "_blank"));

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreateHtmlImageProfile()).Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"https://example.com/docs\"", markdown, StringComparison.Ordinal);
        Assert.Contains("<img src=\"https://example.com/logo.png\" alt=\"Logo\" title=\"Example\" width=\"256\" height=\"128\"", markdown, StringComparison.Ordinal);
        Assert.Contains("target=\"_blank\"", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("{width=", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_Block_Render_Extension_Can_Use_Public_Write_Context() {
        var doc = MarkdownDoc.Create()
            .H1("Intro")
            .H2("Child")
            .TocHere(options => {
                options.IncludeTitle = false;
                options.MinLevel = 2;
                options.MaxLevel = 6;
            });

        var options = new MarkdownWriteOptions();
        options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "toc-context",
            typeof(TocBlock),
            static (block, context) => {
                if (block is not TocBlock toc) {
                    return null;
                }

                var blockIndex = context.GetBlockIndex(toc);
                var anchor = context.GetHeadingAnchor(context.Blocks[0]);
                var entries = context.BuildTocEntries(blockIndex, new TocOptions {
                    IncludeTitle = false,
                    MinLevel = 2,
                    MaxLevel = 6
                });
                return $"<!-- toc-index:{blockIndex}; anchor:{anchor}; entries:{string.Join(",", entries.Select(entry => entry.Anchor))} -->";
            }));

        var markdown = doc.ToMarkdown(options).Replace("\r\n", "\n");

        Assert.Contains("<!-- toc-index:2; anchor:intro; entries:intro,child -->", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- [Child](#child)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_Block_Render_Extension_Can_Read_Final_Syntax_Node_And_Source_Slices() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var readerOptions = new MarkdownReaderOptions { PreserveTrivia = true };
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions).Document;
        MarkdownSyntaxNode? seenSyntax = null;
        MarkdownSourceSlice normalizedSlice = default;
        MarkdownSourceSlice originalSlice = default;
        var normalizedOk = false;
        var originalOk = false;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "quote-source-aware",
            typeof(QuoteBlock),
            (block, context) => {
                if (block is not QuoteBlock quote) {
                    return null;
                }

                seenSyntax = context.FindSyntaxNode(quote);
                normalizedOk = context.TryCreateSourceSlice(quote, out normalizedSlice);
                originalOk = context.TryCreateOriginalSourceSlice(quote, out originalSlice);
                return "> source-aware";
            }));

        var rendered = document.ToMarkdown(options);

        Assert.Contains("> source-aware", rendered, StringComparison.Ordinal);
        Assert.Contains("Tail", rendered, StringComparison.Ordinal);
        Assert.NotNull(seenSyntax);
        Assert.Equal(MarkdownSyntaxKind.Quote, seenSyntax!.Kind);
        Assert.True(normalizedOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, normalizedSlice.TextKind);
        Assert.Equal("> Alpha\n> Beta", normalizedSlice.Text);
        Assert.True(originalOk);
        Assert.Equal(MarkdownSourceTextKind.Original, originalSlice.TextKind);
        Assert.Equal("> Alpha\r\n> Beta", originalSlice.Text);
    }

    [Fact]
    public void Markdown_Syntax_Block_Render_Extension_Can_Read_Final_Syntax_Node_And_Source_Slices() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var readerOptions = new MarkdownReaderOptions { PreserveTrivia = true };
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions).Document;
        MarkdownSyntaxNode? seenSyntax = null;
        MarkdownSourceSlice originalSlice = default;
        var originalOk = false;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockMarkdownRenderExtension.CreateContextual(
            "quote-syntax-source-aware",
            MarkdownSyntaxKind.Quote,
            (block, syntaxNode, context) => {
                seenSyntax = syntaxNode;
                originalOk = context.TryCreateOriginalSourceSlice(syntaxNode, out originalSlice);
                return $"<!-- syntax:{syntaxNode.Kind}; source:{originalSlice.Text.Replace("\r\n", "|")} -->";
            }));

        var rendered = document.ToMarkdown(options);

        Assert.Contains("<!-- syntax:Quote; source:> Alpha|> Beta -->", rendered, StringComparison.Ordinal);
        Assert.Contains("Tail", rendered, StringComparison.Ordinal);
        Assert.NotNull(seenSyntax);
        Assert.Equal(MarkdownSyntaxKind.Quote, seenSyntax!.Kind);
        Assert.True(originalOk);
        Assert.Equal(MarkdownSourceTextKind.Original, originalSlice.TextKind);
    }

    [Fact]
    public void Markdown_Syntax_Inline_Render_Extension_Runs_Before_Type_Extension_And_Can_Read_Source_Slice() {
        const string markdown = "Use `code` now.";
        var readerOptions = new MarkdownReaderOptions { PreserveTrivia = true };
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions).Document;
        MarkdownSourceSlice sourceSlice = default;
        var sourceOk = false;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.InlineRenderExtensions.Add(MarkdownInlineMarkdownRenderExtension.CreateContextual(
            "code-type",
            typeof(CodeSpanInline),
            static (_, _) => "`type`"));
        options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineMarkdownRenderExtension.CreateContextual(
            "code-syntax",
            MarkdownSyntaxKind.InlineCodeSpan,
            (inline, syntaxNode, context) => {
                sourceOk = context.TryCreateSourceSlice(syntaxNode, out sourceSlice);
                return $"`{syntaxNode.Kind}:{sourceSlice.Text}`";
            }));

        var rendered = document.ToMarkdown(options).Trim();

        Assert.Equal("Use `InlineCodeSpan:`code`` now.", rendered);
        Assert.True(sourceOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, sourceSlice.TextKind);
        Assert.Equal("`code`", sourceSlice.Text);
    }

    [Fact]
    public void Html_Syntax_Block_Render_Extension_Can_Read_Final_Syntax_Node_And_Source_Slices() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var readerOptions = new MarkdownReaderOptions { PreserveTrivia = true };
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions).Document;
        MarkdownSourceSlice originalSlice = default;
        var originalOk = false;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockHtmlRenderExtension.CreateContextual(
            "quote-syntax-html",
            MarkdownSyntaxKind.Quote,
            (block, syntaxNode, context) => {
                originalOk = context.TryCreateOriginalSourceSlice(syntaxNode, out originalSlice);
                return $"<aside data-kind=\"{syntaxNode.Kind}\" data-source=\"{System.Net.WebUtility.HtmlEncode(originalSlice.Text.Replace("\r\n", "|"))}\"></aside>";
            }));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<aside data-kind=\"Quote\" data-source=\"&gt; Alpha|&gt; Beta\"></aside>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Tail</p>", html, StringComparison.Ordinal);
        Assert.True(originalOk);
        Assert.Equal(MarkdownSourceTextKind.Original, originalSlice.TextKind);
    }

    [Fact]
    public void Html_Syntax_Inline_Render_Extension_Runs_Before_Type_Extension_And_Can_Read_Source_Slice() {
        const string markdown = "Use `code` now.";
        var readerOptions = new MarkdownReaderOptions { PreserveTrivia = true };
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions).Document;
        MarkdownSourceSlice sourceSlice = default;
        var sourceOk = false;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.InlineRenderExtensions.Add(MarkdownInlineHtmlRenderExtension.CreateContextual(
            "code-type",
            typeof(CodeSpanInline),
            static (_, _) => "<code>type</code>"));
        options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineHtmlRenderExtension.CreateContextual(
            "code-syntax",
            MarkdownSyntaxKind.InlineCodeSpan,
            (inline, syntaxNode, context) => {
                sourceOk = context.TryCreateSourceSlice(syntaxNode, out sourceSlice);
                return $"<kbd data-kind=\"{syntaxNode.Kind}\">{System.Net.WebUtility.HtmlEncode(sourceSlice.Text)}</kbd>";
            }));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<p>Use <kbd data-kind=\"InlineCodeSpan\">`code`</kbd> now.</p>", html, StringComparison.Ordinal);
        Assert.True(sourceOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, sourceSlice.TextKind);
        Assert.Equal("`code`", sourceSlice.Text);
    }

    [Fact]
    public void Markdown_Syntax_Block_Render_Extension_Applies_To_Nested_Quote_Children() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockMarkdownRenderExtension.CreateContextual(
            "nested-paragraph-syntax-markdown",
            MarkdownSyntaxKind.Paragraph,
            static (block, syntaxNode, context) =>
                syntaxNode.Ancestors().Any(parent => parent.Kind == MarkdownSyntaxKind.Quote)
                    ? $"nested:{syntaxNode.Parent?.Kind}"
                    : null));

        var rendered = document.ToMarkdown(options);

        Assert.Contains("> nested:Quote", rendered, StringComparison.Ordinal);
        Assert.Contains("Tail", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_Type_Block_Render_Extension_Applies_To_Nested_Quote_Children() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "nested-paragraph-type-markdown",
            typeof(ParagraphBlock),
            static (block, context) =>
                block is ParagraphBlock paragraph && context.GetBlockIndex(paragraph) < 0
                    ? "nested:type"
                    : null));

        var rendered = document.ToMarkdown(options);

        Assert.Contains("> nested:type", rendered, StringComparison.Ordinal);
        Assert.Contains("Tail", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_Syntax_Block_Render_Extension_Applies_To_Nested_Quote_Children() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockHtmlRenderExtension.CreateContextual(
            "nested-paragraph-syntax-html",
            MarkdownSyntaxKind.Paragraph,
            static (block, syntaxNode, context) =>
                syntaxNode.Ancestors().Any(parent => parent.Kind == MarkdownSyntaxKind.Quote)
                    ? $"<p data-nested=\"{syntaxNode.Parent?.Kind}\">syntax</p>"
                    : null));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<blockquote><p data-nested=\"Quote\">syntax</p></blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Tail</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_Type_Block_Render_Extension_Applies_To_Nested_Quote_Children() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.BlockRenderExtensions.Add(MarkdownBlockHtmlRenderExtension.CreateContextual(
            "nested-paragraph-type-html",
            typeof(ParagraphBlock),
            static (block, context) =>
                block is ParagraphBlock paragraph && context.GetBlockIndex(paragraph) < 0
                    ? "<p data-nested=\"type\">html</p>"
                    : null));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<blockquote><p data-nested=\"type\">html</p></blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Tail</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_Write_Context_Can_Render_Custom_Container_Children_Through_Overrides() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockMarkdownRenderExtension.CreateContextual(
            "context-child-paragraph-markdown",
            MarkdownSyntaxKind.Paragraph,
            static (block, syntaxNode, context) =>
                syntaxNode.Ancestors().Any(parent => parent.Kind == MarkdownSyntaxKind.Quote)
                    ? "child-through-context"
                    : null));
        options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "quote-context-container",
            typeof(QuoteBlock),
            static (block, context) => block is QuoteBlock quote
                ? string.Join("\n", quote.ChildBlocks.Select(context.RenderBlock))
                : null));

        var rendered = document.ToMarkdown(options);

        Assert.Contains("child-through-context", rendered, StringComparison.Ordinal);
        Assert.Contains("Tail", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_Render_Context_Can_Render_Custom_Container_Children_Through_Overrides() {
        const string markdown = "> Alpha\r\n> Beta\r\n\r\nTail\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockHtmlRenderExtension.CreateContextual(
            "context-child-paragraph-html",
            MarkdownSyntaxKind.Paragraph,
            static (block, syntaxNode, context) =>
                syntaxNode.Ancestors().Any(parent => parent.Kind == MarkdownSyntaxKind.Quote)
                    ? "<p data-context-child=\"true\">child</p>"
                    : null));
        options.BlockRenderExtensions.Add(MarkdownBlockHtmlRenderExtension.CreateContextual(
            "quote-context-container",
            typeof(QuoteBlock),
            static (block, context) => block is QuoteBlock quote
                ? "<aside>" + string.Concat(quote.ChildBlocks.Select(context.RenderBlock)) + "</aside>"
                : null));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<aside><p data-context-child=\"true\">child</p></aside>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Tail</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_Writer_Lengthens_Custom_Container_Fence_Around_Nested_Custom_Container() {
        var document = MarkdownDoc.Create().Add(new CustomContainerBlock(
            "outer",
            new IMarkdownBlock[] {
                new CustomContainerBlock(
                    "inner",
                    new IMarkdownBlock[] {
                        new ParagraphBlock(new InlineSequence().Text("hello"))
                    })
            }));

        var written = document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd();

        Assert.Equal(":::: outer\n::: inner\nhello\n:::\n::::", written);

        var readerOptions = MarkdownReaderOptions.CreatePortableProfile();
        readerOptions.CustomContainers = true;
        var reparsed = MarkdownReader.Parse(written, readerOptions);

        var outer = Assert.IsType<CustomContainerBlock>(Assert.Single(reparsed.Blocks));
        var inner = Assert.IsType<CustomContainerBlock>(Assert.Single(outer.ChildBlocks));
        Assert.Equal("outer", outer.Name);
        Assert.Equal("inner", inner.Name);
        Assert.Equal("hello", Assert.IsType<ParagraphBlock>(Assert.Single(inner.ChildBlocks)).Inlines.RenderMarkdown());
    }

    [Fact]
    public void Markdown_Writer_Lengthens_Custom_Container_Fence_Around_List_Item_Custom_Container() {
        var item = ListItem.Text("item");
        item.Children.Add(new CustomContainerBlock(
            "inner",
            new IMarkdownBlock[] {
                new ParagraphBlock(new InlineSequence().Text("hello"))
            }));
        var list = new UnorderedListBlock();
        list.Items.Add(item);
        var document = MarkdownDoc.Create().Add(new CustomContainerBlock(
            "outer",
            new IMarkdownBlock[] { list }));

        var written = document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd();

        Assert.Equal(":::: outer\n- item\n  ::: inner\n  hello\n  :::\n::::", written);

        var readerOptions = MarkdownReaderOptions.CreatePortableProfile();
        readerOptions.CustomContainers = true;
        var reparsed = MarkdownReader.Parse(written, readerOptions);

        var outer = Assert.IsType<CustomContainerBlock>(Assert.Single(reparsed.Blocks));
        var reparsedList = Assert.IsType<UnorderedListBlock>(Assert.Single(outer.ChildBlocks));
        var reparsedItem = Assert.Single(reparsedList.Items);
        var inner = Assert.IsType<CustomContainerBlock>(Assert.Single(reparsedItem.Children));
        Assert.Equal("outer", outer.Name);
        Assert.Equal("inner", inner.Name);
        Assert.Equal("hello", Assert.IsType<ParagraphBlock>(Assert.Single(inner.ChildBlocks)).Inlines.RenderMarkdown());
    }

    [Fact]
    public void Markdown_Writer_Keeps_List_Item_First_Line_Definition_Like_Text_Literal() {
        var list = new UnorderedListBlock();
        var labelItem = ListItem.Text("Formula:");
        labelItem.AdditionalParagraphs.Add(new InlineSequence().Text("Unsupported Word content: equation"));
        list.Items.Add(labelItem);
        list.Items.Add(ListItem.Text("[ref]: https://example.com"));

        var written = MarkdownDoc.Create()
            .Add(list)
            .ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" })
            .TrimEnd();

        Assert.Equal("- Formula:\n\n  Unsupported Word content: equation\n- \\[ref\\]&#58; https://example.com", written);
    }

    [Fact]
    public void Html_Syntax_Block_Render_Extension_Applies_To_Custom_Container_In_Tight_List_Item() {
        const string markdown = """
- item
  ::: note
  hello
  :::
""";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true,
            CustomContainers = true
        }).Document;
        MarkdownSyntaxNode? seenSyntax = null;
        MarkdownSourceSlice sourceSlice = default;
        var sourceOk = false;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockHtmlRenderExtension.CreateContextual(
            "tight-list-custom-container-html",
            MarkdownSyntaxKind.CustomContainer,
            (block, syntaxNode, context) => {
                seenSyntax = syntaxNode;
                sourceOk = context.TryCreateSourceSlice(syntaxNode, out sourceSlice);
                var name = block is CustomContainerBlock container ? container.Name : string.Empty;
                return $"<aside data-name=\"{context.EncodeAttributeValue(name)}\">custom</aside>";
            }));

        var html = document.ToHtmlFragment(options).Replace("\r\n", "\n");

        Assert.Contains("<ul><li>item <aside data-name=\"note\">custom</aside></li></ul>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<div class=\"note\">", html, StringComparison.Ordinal);
        Assert.NotNull(seenSyntax);
        Assert.Equal(MarkdownSyntaxKind.CustomContainer, seenSyntax!.Kind);
        Assert.True(sourceOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, sourceSlice.TextKind);
        Assert.Equal("::: note\n  hello\n  :::", sourceSlice.Text);
    }

    [Fact]
    public void Html_Syntax_Block_Render_Extension_Applies_To_Custom_Container_Tight_List_Children() {
        const string markdown = """
- item
  ::: note
  hello
  :::
""";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true,
            CustomContainers = true
        }).Document;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockHtmlRenderExtension.CreateContextual(
            "tight-list-custom-container-child-html",
            MarkdownSyntaxKind.Paragraph,
            static (block, syntaxNode, context) =>
                syntaxNode.Ancestors().Any(parent => parent.Kind == MarkdownSyntaxKind.CustomContainer)
                    ? "<span data-container-child=\"true\">child</span>"
                    : null));

        var html = document.ToHtmlFragment(options).Replace("\r\n", "\n");

        Assert.Contains("<ul><li>item <div class=\"note\"><span data-container-child=\"true\">child</span></div></li></ul>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<p>hello</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_Block_Render_Extension_Can_Create_Source_Slices_From_Token_Source_Spans() {
        const string markdown = "> [!TIP] Heads up\r\n> Body\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;
        MarkdownSourceSlice openingMarkerSlice = default;
        MarkdownSourceSlice kindSlice = default;
        MarkdownSourceSlice closingMarkerSlice = default;
        MarkdownSourceSlice titleSlice = default;
        var openingMarkerOk = false;
        var kindOk = false;
        var closingMarkerOk = false;
        var titleOk = false;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "callout-token-source-markdown",
            typeof(CalloutBlock),
            (block, context) => {
                if (block is not CalloutBlock callout
                    || !callout.OpeningMarkerSourceSpan.HasValue
                    || !callout.KindSourceSpan.HasValue
                    || !callout.ClosingMarkerSourceSpan.HasValue
                    || !callout.TitleSourceSpan.HasValue) {
                    return null;
                }

                openingMarkerOk = context.TryCreateOriginalSourceSlice(callout.OpeningMarkerSourceSpan.Value, out openingMarkerSlice);
                kindOk = context.TryCreateOriginalSourceSlice(callout.KindSourceSpan.Value, out kindSlice);
                closingMarkerOk = context.TryCreateOriginalSourceSlice(callout.ClosingMarkerSourceSpan.Value, out closingMarkerSlice);
                titleOk = context.TryCreateOriginalSourceSlice(callout.TitleSourceSpan.Value, out titleSlice);
                return $"<!-- open:{openingMarkerSlice.Text}; kind:{kindSlice.Text}; close:{closingMarkerSlice.Text}; title:{titleSlice.Text} -->";
            }));

        var rendered = document.ToMarkdown(options);

        Assert.Contains("<!-- open:[!; kind:TIP; close:]; title:Heads up -->", rendered, StringComparison.Ordinal);
        Assert.True(openingMarkerOk);
        Assert.True(kindOk);
        Assert.True(closingMarkerOk);
        Assert.True(titleOk);
        Assert.Equal(MarkdownSourceTextKind.Original, openingMarkerSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, kindSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, closingMarkerSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, titleSlice.TextKind);
    }

    [Fact]
    public void Html_Block_Render_Extension_Can_Create_Source_Slices_From_Token_Source_Spans() {
        const string markdown = "> [!TIP] Heads up\r\n> Body\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;
        MarkdownSourceSlice openingMarkerSlice = default;
        MarkdownSourceSlice kindSlice = default;
        MarkdownSourceSlice closingMarkerSlice = default;
        MarkdownSourceSlice titleSlice = default;
        var openingMarkerOk = false;
        var kindOk = false;
        var closingMarkerOk = false;
        var titleOk = false;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.BlockRenderExtensions.Add(MarkdownBlockHtmlRenderExtension.CreateContextual(
            "callout-token-source-html",
            typeof(CalloutBlock),
            (block, context) => {
                if (block is not CalloutBlock callout
                    || !callout.OpeningMarkerSourceSpan.HasValue
                    || !callout.KindSourceSpan.HasValue
                    || !callout.ClosingMarkerSourceSpan.HasValue
                    || !callout.TitleSourceSpan.HasValue) {
                    return null;
                }

                openingMarkerOk = context.TryCreateOriginalSourceSlice(callout.OpeningMarkerSourceSpan.Value, out openingMarkerSlice);
                kindOk = context.TryCreateOriginalSourceSlice(callout.KindSourceSpan.Value, out kindSlice);
                closingMarkerOk = context.TryCreateOriginalSourceSlice(callout.ClosingMarkerSourceSpan.Value, out closingMarkerSlice);
                titleOk = context.TryCreateOriginalSourceSlice(callout.TitleSourceSpan.Value, out titleSlice);
                return $"<aside data-open-token=\"{System.Net.WebUtility.HtmlEncode(openingMarkerSlice.Text)}\" data-kind-token=\"{kindSlice.Text}\" data-close-token=\"{System.Net.WebUtility.HtmlEncode(closingMarkerSlice.Text)}\" data-title-token=\"{System.Net.WebUtility.HtmlEncode(titleSlice.Text)}\"></aside>";
            }));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<aside data-open-token=\"[!\" data-kind-token=\"TIP\" data-close-token=\"]\" data-title-token=\"Heads up\"></aside>", html, StringComparison.Ordinal);
        Assert.True(openingMarkerOk);
        Assert.True(kindOk);
        Assert.True(closingMarkerOk);
        Assert.True(titleOk);
        Assert.Equal(MarkdownSourceTextKind.Original, openingMarkerSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, kindSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, closingMarkerSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, titleSlice.TextKind);
    }

    [Fact]
    public void Html_Block_Render_Extension_Can_Use_Active_NonAscii_Encoding_Helpers() {
        const string markdown = "> åAlpha\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;
        MarkdownSourceSlice originalSlice = default;
        var originalOk = false;

        var options = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        options.BlockRenderExtensions.Add(MarkdownBlockHtmlRenderExtension.CreateContextual(
            "quote-policy-helpers",
            typeof(QuoteBlock),
            (block, context) => {
                originalOk = context.TryCreateOriginalSourceSlice(block, out originalSlice);
                return "<aside data-source=\"" + context.EncodeAttributeValue(originalSlice.Text) + "\">"
                    + context.EncodeText("åbody < &")
                    + "</aside>";
            }));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<aside data-source=\"&gt; åAlpha\">åbody &lt; &amp;</aside>", html, StringComparison.Ordinal);
        Assert.True(originalOk);
        Assert.Equal(MarkdownSourceTextKind.Original, originalSlice.TextKind);
    }

    [Fact]
    public void Block_Render_Extensions_Report_Original_Source_Slice_Failure_Reasons() {
        const string markdown = "> Alpha\r\n";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown).Document;
        MarkdownOriginalSourceSliceFailureReason markdownReason = MarkdownOriginalSourceSliceFailureReason.None;
        MarkdownOriginalSourceSliceFailureReason htmlReason = MarkdownOriginalSourceSliceFailureReason.None;

        var writeOptions = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        writeOptions.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "quote-original-failure-markdown",
            typeof(QuoteBlock),
            (block, context) => {
                context.TryCreateOriginalSourceSlice(block, out _, out markdownReason);
                return "> markdown-reason:" + markdownReason;
            }));

        var rendered = document.ToMarkdown(writeOptions);

        var htmlOptions = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        htmlOptions.BlockRenderExtensions.Add(MarkdownBlockHtmlRenderExtension.CreateContextual(
            "quote-original-failure-html",
            typeof(QuoteBlock),
            (block, context) => {
                context.TryCreateOriginalSourceSlice(block, out _, out htmlReason);
                return "<aside data-reason=\"" + htmlReason + "\"></aside>";
            }));

        var html = document.ToHtmlFragment(htmlOptions);

        Assert.Contains("> markdown-reason:OriginalMarkdownNotPreserved", rendered, StringComparison.Ordinal);
        Assert.Contains("<aside data-reason=\"OriginalMarkdownNotPreserved\"></aside>", html, StringComparison.Ordinal);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, markdownReason);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, htmlReason);
    }

    [Fact]
    public void Inline_Render_Extensions_Report_Original_Source_Slice_Failure_Reasons() {
        const string markdown = "Use `code` now.";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown).Document;
        MarkdownOriginalSourceSliceFailureReason markdownReason = MarkdownOriginalSourceSliceFailureReason.None;
        MarkdownOriginalSourceSliceFailureReason htmlReason = MarkdownOriginalSourceSliceFailureReason.None;

        var writeOptions = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        writeOptions.InlineRenderExtensions.Add(MarkdownInlineMarkdownRenderExtension.CreateContextual(
            "code-original-failure-markdown",
            typeof(CodeSpanInline),
            (inline, context) => {
                context.TryCreateOriginalSourceSlice(inline, out _, out markdownReason);
                return "`" + markdownReason + "`";
            }));

        var rendered = document.ToMarkdown(writeOptions);

        var htmlOptions = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        htmlOptions.InlineRenderExtensions.Add(MarkdownInlineHtmlRenderExtension.CreateContextual(
            "code-original-failure-html",
            typeof(CodeSpanInline),
            (inline, context) => {
                context.TryCreateOriginalSourceSlice(inline, out _, out htmlReason);
                return "<kbd data-reason=\"" + htmlReason + "\">code</kbd>";
            }));

        var html = document.ToHtmlFragment(htmlOptions);

        Assert.Equal("Use `OriginalMarkdownNotPreserved` now.\n", rendered);
        Assert.Contains("<p>Use <kbd data-reason=\"OriginalMarkdownNotPreserved\">code</kbd> now.</p>", html, StringComparison.Ordinal);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, markdownReason);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, htmlReason);
    }

    [Fact]
    public void Markdown_Inline_Render_Extension_Can_Create_Source_Slices_From_Metadata_Source_Spans() {
        const string markdown = "Go [there](https://example.com \"Example\") now.";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;
        MarkdownSourceSlice targetSlice = default;
        MarkdownSourceSlice titleSlice = default;
        var targetOk = false;
        var titleOk = false;

        var options = new MarkdownWriteOptions { OutputLineEnding = "\n" };
        options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineMarkdownRenderExtension.CreateContextual(
            "link-token-source-markdown",
            MarkdownSyntaxKind.InlineLink,
            (inline, syntaxNode, context) => {
                var targetNode = syntaxNode.Children.FirstOrDefault(child => child.Kind == MarkdownSyntaxKind.InlineLinkTarget);
                var titleNode = syntaxNode.Children.FirstOrDefault(child => child.Kind == MarkdownSyntaxKind.InlineLinkTitle);
                if (targetNode?.SourceSpan == null || titleNode?.SourceSpan == null) {
                    return null;
                }

                targetOk = context.TryCreateSourceSlice(targetNode.SourceSpan.Value, out targetSlice);
                titleOk = context.TryCreateSourceSlice(titleNode.SourceSpan.Value, out titleSlice);
                return $"[{targetSlice.Text}|{titleSlice.Text}]";
            }));

        var rendered = document.ToMarkdown(options).Trim();

        Assert.Equal("Go [https://example.com|Example] now.", rendered);
        Assert.True(targetOk);
        Assert.True(titleOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, targetSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Normalized, titleSlice.TextKind);
    }

    [Fact]
    public void Html_Inline_Render_Extension_Can_Use_Active_NonAscii_Encoding_Helpers() {
        const string markdown = "Use `åcode` now.";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;
        MarkdownSourceSlice sourceSlice = default;
        var sourceOk = false;

        var options = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineHtmlRenderExtension.CreateContextual(
            "inline-policy-helpers",
            MarkdownSyntaxKind.InlineCodeSpan,
            (inline, syntaxNode, context) => {
                sourceOk = context.TryCreateSourceSlice(syntaxNode, out sourceSlice);
                return "<kbd data-source=\"" + context.EncodeAttributeValue(sourceSlice.Text) + "\">"
                    + context.EncodeText("åinline < &")
                    + "</kbd>";
            }));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<p>Use <kbd data-source=\"`åcode`\">åinline &lt; &amp;</kbd> now.</p>", html, StringComparison.Ordinal);
        Assert.True(sourceOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, sourceSlice.TextKind);
    }

    [Fact]
    public void Html_Inline_Render_Extension_Can_Create_Source_Slices_From_Metadata_Source_Spans() {
        const string markdown = "Go [there](https://example.com \"Example\") now.";
        var document = MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions { PreserveTrivia = true }).Document;
        MarkdownSourceSlice targetSlice = default;
        MarkdownSourceSlice titleSlice = default;
        var targetOk = false;
        var titleOk = false;

        var options = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };
        options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineHtmlRenderExtension.CreateContextual(
            "link-token-source-html",
            MarkdownSyntaxKind.InlineLink,
            (inline, syntaxNode, context) => {
                var targetNode = syntaxNode.Children.FirstOrDefault(child => child.Kind == MarkdownSyntaxKind.InlineLinkTarget);
                var titleNode = syntaxNode.Children.FirstOrDefault(child => child.Kind == MarkdownSyntaxKind.InlineLinkTitle);
                if (targetNode?.SourceSpan == null || titleNode?.SourceSpan == null) {
                    return null;
                }

                targetOk = context.TryCreateSourceSlice(targetNode.SourceSpan.Value, out targetSlice);
                titleOk = context.TryCreateSourceSlice(titleNode.SourceSpan.Value, out titleSlice);
                return $"<a data-target-token=\"{System.Net.WebUtility.HtmlEncode(targetSlice.Text)}\" data-title-token=\"{System.Net.WebUtility.HtmlEncode(titleSlice.Text)}\">token</a>";
            }));

        var html = document.ToHtmlFragment(options);

        Assert.Contains("<p>Go <a data-target-token=\"https://example.com\" data-title-token=\"Example\">token</a> now.</p>", html, StringComparison.Ordinal);
        Assert.True(targetOk);
        Assert.True(titleOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, targetSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Normalized, titleSlice.TextKind);
    }

    [Fact]
    public void Markdown_Block_Render_Extension_Legacy_Constructor_Still_Uses_Options_And_Applies() {
        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""");
        var options = new MarkdownWriteOptions();
        options.BlockRenderExtensions.Add(new MarkdownBlockMarkdownRenderExtension(
            "callout-legacy",
            typeof(CalloutBlock),
            static (block, writerOptions) => {
                if (block is not CalloutBlock callout) {
                    return null;
                }

                return $"<!-- mode:{writerOptions.ImageRenderingMode}; kind:{callout.Kind} -->";
            }));

        var markdown = doc.ToMarkdown(options).Replace("\r\n", "\n");

        Assert.Contains("<!-- mode:RichMarkdown; kind:note -->", markdown, StringComparison.Ordinal);
    }
}
