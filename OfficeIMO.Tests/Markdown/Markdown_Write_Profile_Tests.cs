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
