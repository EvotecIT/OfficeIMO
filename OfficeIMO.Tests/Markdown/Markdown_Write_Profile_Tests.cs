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
