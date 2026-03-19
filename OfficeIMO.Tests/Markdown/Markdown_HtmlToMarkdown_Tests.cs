using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.SamplePlugin;
using Xunit;
using MarkdownRendererShell = OfficeIMO.MarkdownRenderer.MarkdownRenderer;
using System.IO;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlToMarkdownTests {
    [Fact]
    public void HtmlToMarkdown_Converter_UsesDefaultOptionsWhenNull() {
        var converter = new HtmlToMarkdownConverter();

        string markdown = converter.Convert("<p>Hello</p>", options: null);
        MarkdownDoc document = converter.ConvertToDocument("<p>Hello</p>", options: null);

        Assert.Contains("Hello", markdown, StringComparison.Ordinal);
        Assert.Single(document.Blocks);
        Assert.IsType<ParagraphBlock>(document.Blocks[0]);
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsCommonDocumentBlocks() {
        string html = "<html><body><h1>Hello</h1><p>A <strong>bold</strong> <a href=\"https://example.com\">link</a>.</p><ul><li>One</li><li>Two</li></ul></body></html>";

        string markdown = html.ToMarkdown();

        Assert.Contains("# Hello", markdown, StringComparison.Ordinal);
        Assert.Contains("**bold**", markdown, StringComparison.Ordinal);
        Assert.Contains("[link](https://example.com)", markdown, StringComparison.Ordinal);
        Assert.Contains("- One", markdown, StringComparison.Ordinal);
        Assert.Contains("- Two", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Trims_StrongBoundaryWhitespace_Before_InlineParsing() {
        string html = "<p><strong> LDAP/Kerberos health on all DCs </strong> next</p>";

        MarkdownDoc document = html.LoadFromHtml();
        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>LDAP/Kerberos health on all DCs</strong> next", renderedHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("** LDAP/Kerberos health on all DCs **", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Convert_Uses_Configured_MarkdownWriteOptions() {
        var options = new HtmlToMarkdownOptions {
            MarkdownWriteOptions = new MarkdownWriteOptions()
        };
        options.MarkdownWriteOptions.BlockRenderExtensions.Add(new MarkdownBlockMarkdownRenderExtension(
            "Test.Paragraph.Override",
            typeof(ParagraphBlock),
            (block, _) => block is ParagraphBlock ? "PARAGRAPH-OVERRIDE" : null));

        string markdown = "<p>Hello</p>".ToMarkdown(options);

        Assert.Equal("PARAGRAPH-OVERRIDE", markdown.Trim());
    }

    [Fact]
    public void HtmlToMarkdown_LoadFromHtml_ProducesTypedBlocks() {
        string html = "<html><body><h2>Section</h2><blockquote><p>Quoted</p></blockquote><details open><summary>More</summary><p>Hidden text</p></details></body></html>";

        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Level == 2 && heading.Text == "Section");
        Assert.Contains(document.Blocks, block => block is QuoteBlock);
        Assert.Contains(document.Blocks, block => block is DetailsBlock details && details.Open);
    }

    [Fact]
    public void HtmlToMarkdown_Enforces_MaxInputCharacters() {
        var options = new HtmlToMarkdownOptions {
            MaxInputCharacters = 12
        };

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => "<p>0123456789</p>".LoadFromHtml(options));
        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsHtmlFragmentWithoutBodyWrapper() {
        string html = "<h2>Fragment</h2><p>Body</p>";

        MarkdownDoc document = html.LoadFromHtml();

        Assert.Collection(document.Blocks,
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_ResolvesRelativeLinksWithBaseUri() {
        string html = "<p><a href=\"guide/start\">Docs</a></p>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/")
        });

        Assert.Contains("[Docs](https://example.com/docs/guide/start)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_UsesDocumentBaseHrefToResolveRelativeLinks() {
        const string html = """
<html>
  <head><base href="https://cdn.example.com/docs/" /></head>
  <body><p><a href="guide/start">Docs</a></p></body>
</html>
""";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/app/")
        });

        Assert.Contains("[Docs](https://cdn.example.com/docs/guide/start)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_UsesRelativeDocumentBaseHrefAgainstProvidedPageUri() {
        const string html = """
<html>
  <head><base href="/assets/" /></head>
  <body><figure><img src="images/demo.png" alt="Demo" /></figure></body>
</html>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/start/page.html")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/assets/images/demo.png", image.Path);
        Assert.Equal("Demo", image.Alt);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesInlineLinkHtmlMetadataInTypedAst() {
        const string html = """
<p><a href="/docs/hero" title="Hero docs" target="_blank" rel="nofollow sponsored">Read more</a></p>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var link = Assert.IsType<LinkInline>(Assert.Single(paragraph.Inlines.Nodes));
        Assert.Equal("Read more", link.Text);
        Assert.Equal("https://example.com/docs/hero", link.Url);
        Assert.Equal("Hero docs", link.Title);
        Assert.Equal("_blank", link.LinkTarget);
        Assert.Equal("nofollow sponsored", link.LinkRel);

        string markdown = document.ToMarkdown();
        Assert.Contains("[Read more](https://example.com/docs/hero \"Hero docs\")", markdown, StringComparison.Ordinal);

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<a href=\"https://example.com/docs/hero\" title=\"Hero docs\" target=\"_blank\" rel=\"", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("nofollow", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("sponsored", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("noopener", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("noreferrer", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlocks_WhenRequested() {
        string html = "<custom-widget data-name=\"demo\">Hello</custom-widget>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Contains("<custom-widget", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedInlineHtml_WhenRequested() {
        string html = "<p>Hello <custom-inline data-name=\"demo\">world</custom-inline></p>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedInlineHtml = true
        });

        Assert.Contains("<custom-inline", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadFromHtml_PreservesUnsupportedInlineHtmlInAst() {
        string html = "<p>Hello <custom-inline data-name=\"demo\">world</custom-inline></p>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedInlineHtml = true
        });

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        Assert.Contains(paragraph.Inlines.Nodes, inline => inline is HtmlRawInline raw && raw.Html.Contains("<custom-inline", StringComparison.Ordinal));

        string markdown = document.ToMarkdown();
        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<custom-inline", markdown, StringComparison.Ordinal);
        Assert.Contains("<custom-inline", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadFromHtml_PreservesExtendedInlineHtmlTagsAsTypedNodes() {
        const string html = "<p>Before <q>quoted</q> H<sub>2</sub>O <ins>inserted</ins> x<sup>2</sup></p>";

        MarkdownDoc document = html.LoadFromHtml();
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.Contains(paragraph.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "q");
        Assert.Contains(paragraph.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "sub");
        Assert.Contains(paragraph.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "ins");
        Assert.Contains(paragraph.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "sup");

        string markdown = document.ToMarkdown();
        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<q>quoted</q>", markdown, StringComparison.Ordinal);
        Assert.Contains("<sub>2</sub>", markdown, StringComparison.Ordinal);
        Assert.Contains("<ins>inserted</ins>", markdown, StringComparison.Ordinal);
        Assert.Contains("<sup>2</sup>", markdown, StringComparison.Ordinal);
        Assert.Contains("<q>quoted</q>", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<sub>2</sub>", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<ins>inserted</ins>", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<sup>2</sup>", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdownOptions_Clone_CopiesMarkdownWriteOptionsCollections() {
        var options = new HtmlToMarkdownOptions {
            MarkdownWriteOptions = new MarkdownWriteOptions()
        };
        options.MarkdownWriteOptions.BlockRenderExtensions.Add(new MarkdownBlockMarkdownRenderExtension(
            "Original.Extension",
            typeof(ParagraphBlock),
            (block, _) => block is ParagraphBlock ? "ORIGINAL" : null));

        var clone = options.Clone();
        Assert.NotSame(options.MarkdownWriteOptions, clone.MarkdownWriteOptions);
        Assert.Single(options.MarkdownWriteOptions.BlockRenderExtensions);
        Assert.Single(clone.MarkdownWriteOptions!.BlockRenderExtensions);

        clone.MarkdownWriteOptions.BlockRenderExtensions.Add(new MarkdownBlockMarkdownRenderExtension(
            "Clone.Extension",
            typeof(HeadingBlock),
            (block, _) => block is HeadingBlock ? "CLONE" : null));

        Assert.Single(options.MarkdownWriteOptions.BlockRenderExtensions);
        Assert.Equal(2, clone.MarkdownWriteOptions.BlockRenderExtensions.Count);
    }

    [Fact]
    public void HtmlToMarkdownOptions_Clone_CopiesVisualElementRoundTripHints() {
        var options = new HtmlToMarkdownOptions();
        options.DocumentTransforms.Add(SampleMarkdownRenderer.StatusPanelHtmlDocumentTransform);
        options.ElementBlockConverters.Add(SampleMarkdownRenderer.StatusPanelVendorHtmlConverter);
        options.InlineElementConverters.Add(SampleMarkdownRenderer.StatusBadgeInlineConverter);
        options.VisualElementRoundTripHints.Add(new MarkdownVisualElementRoundTripHint(
            "vendor.caption",
            "Vendor caption",
            context => context.CreateBlock(caption: "Caption")));
        options.TryMarkPluginApplied("vendor.visuals");
        options.TryMarkFeaturePackApplied("vendor.visual-pack");

        var clone = options.Clone();

        Assert.Single(options.DocumentTransforms);
        Assert.Single(clone.DocumentTransforms);
        Assert.Same(options.DocumentTransforms[0], clone.DocumentTransforms[0]);
        Assert.Single(options.ElementBlockConverters);
        Assert.Single(clone.ElementBlockConverters);
        Assert.Same(options.ElementBlockConverters[0], clone.ElementBlockConverters[0]);
        Assert.Single(options.InlineElementConverters);
        Assert.Single(clone.InlineElementConverters);
        Assert.Same(options.InlineElementConverters[0], clone.InlineElementConverters[0]);
        Assert.Single(options.VisualElementRoundTripHints);
        Assert.Single(clone.VisualElementRoundTripHints);
        Assert.Same(options.VisualElementRoundTripHints[0], clone.VisualElementRoundTripHints[0]);
        Assert.True(clone.HasPluginId("vendor.visuals"));
        Assert.True(clone.HasFeaturePackId("vendor.visual-pack"));

        clone.DocumentTransforms.Add(new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));
        clone.ElementBlockConverters.Add(new HtmlElementBlockConverter(
            "vendor.secondary-html",
            "Vendor secondary HTML",
            _ => Array.Empty<IMarkdownBlock>()));
        clone.InlineElementConverters.Add(new HtmlInlineElementConverter(
            "vendor.secondary-inline",
            "Vendor secondary inline HTML",
            _ => Array.Empty<IMarkdownInline>()));
        clone.VisualElementRoundTripHints.Add(new MarkdownVisualElementRoundTripHint(
            "vendor.secondary-caption",
            "Vendor caption 2",
            context => context.CreateBlock(caption: "Caption 2")));

        Assert.Single(options.DocumentTransforms);
        Assert.Equal(2, clone.DocumentTransforms.Count);
        Assert.Single(options.ElementBlockConverters);
        Assert.Equal(2, clone.ElementBlockConverters.Count);
        Assert.Single(options.InlineElementConverters);
        Assert.Equal(2, clone.InlineElementConverters.Count);
        Assert.Single(options.VisualElementRoundTripHints);
        Assert.Equal(2, clone.VisualElementRoundTripHints.Count);
    }

    [Fact]
    public void HtmlToMarkdownOptions_Can_Apply_Renderer_Plugin_RoundTrip_Hints_Idempotently() {
        var options = new HtmlToMarkdownOptions();

        options.ApplyPlugin(SampleMarkdownRenderer.StatusPanelPlugin);
        options.ApplyPlugin(SampleMarkdownRenderer.StatusPanelPlugin);

        Assert.True(options.HasPlugin(SampleMarkdownRenderer.StatusPanelPlugin));
        Assert.Single(options.DocumentTransforms);
        Assert.Same(SampleMarkdownRenderer.StatusPanelHtmlDocumentTransform, options.DocumentTransforms[0]);
        Assert.Single(options.ElementBlockConverters);
        Assert.Same(SampleMarkdownRenderer.StatusPanelVendorHtmlConverter, options.ElementBlockConverters[0]);
        Assert.Single(options.InlineElementConverters);
        Assert.Same(SampleMarkdownRenderer.StatusBadgeInlineConverter, options.InlineElementConverters[0]);
        Assert.Single(options.VisualElementRoundTripHints);
        Assert.Equal("sample.status-panel-caption", options.VisualElementRoundTripHints[0].Id);
    }

    [Fact]
    public void HtmlToMarkdownOptions_Can_Apply_Renderer_FeaturePack_RoundTrip_Hints_Idempotently() {
        var options = new HtmlToMarkdownOptions();

        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        Assert.True(options.HasFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack));
        Assert.True(options.HasPlugin(SampleMarkdownRenderer.StatusPanelPlugin));
        Assert.Single(options.DocumentTransforms);
        Assert.Same(SampleMarkdownRenderer.StatusPanelHtmlDocumentTransform, options.DocumentTransforms[0]);
        Assert.Single(options.ElementBlockConverters);
        Assert.Same(SampleMarkdownRenderer.StatusPanelVendorHtmlConverter, options.ElementBlockConverters[0]);
        Assert.Single(options.InlineElementConverters);
        Assert.Same(SampleMarkdownRenderer.StatusBadgeInlineConverter, options.InlineElementConverters[0]);
        Assert.Single(options.VisualElementRoundTripHints);
        Assert.Equal("sample.status-panel-caption", options.VisualElementRoundTripHints[0].Id);
    }

    [Fact]
    public void HtmlToMarkdown_Applies_Sample_StatusBadge_Inline_Converter_To_Vendor_Html() {
        const string html = "<p>Server <span class=\"sample-status-badge\">Healthy</span> now</p>";
        var options = new HtmlToMarkdownOptions();
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        MarkdownDoc document = html.LoadFromHtml(options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var highlight = Assert.IsType<HighlightSequenceInline>(Assert.Single(paragraph.Inlines.Nodes.OfType<HighlightSequenceInline>()));
        Assert.Equal("Healthy", highlight.Inlines.RenderMarkdown());
        Assert.Equal(
            NormalizeMarkdown("Server ==Healthy== now"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsBetweenParagraphs() {
        string html = "<p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Collection(document.Blocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int widgetIndex = markdown.IndexOf("<custom-widget", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(widgetIndex > alphaIndex, "Expected raw custom block after the opening paragraph.");
        Assert.True(omegaIndex > widgetIndex, "Expected trailing paragraph after the custom block.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsInsideListItems() {
        string html = "<ul><li><p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsInsideSections() {
        string html = "<section><p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p></section>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Collection(document.Blocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsInsideDetails() {
        string html = "<details open><summary>More</summary><p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p></details>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        var details = Assert.IsType<DetailsBlock>(Assert.Single(document.Blocks));
        Assert.NotNull(details.Summary);
        Assert.Collection(details.ChildBlocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_CapturesFigureCaptionOnImageBlocks() {
        string html = "<figure><img src=\"/img/demo.png\" alt=\"Demo\" /><figcaption>Example caption</figcaption></figure>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/demo.png", image.Path);
        Assert.Equal("Demo", image.Alt);
        Assert.Equal("Example caption", image.Caption);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesFigureOrderInsideListItems() {
        string html = "<ul><li><p>Alpha</p><figure><img src=\"/img/demo.png\" alt=\"Demo\" /><figcaption>Caption</figcaption></figure><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/img/demo.png", image.Path);
                Assert.Equal("Caption", image.Caption);
            },
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesSupplementalFigureBlocksAroundDirectImage() {
        const string html = """
<figure>
  <p>Lead-in</p>
  <img src="/img/demo.png" alt="Demo" />
  <figcaption>Caption</figcaption>
  <blockquote><p>Quoted note</p></blockquote>
</figure>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Lead-in", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/img/demo.png", image.Path);
                Assert.Equal("Demo", image.Alt);
                Assert.Equal("Caption", image.Caption);
            },
            block => Assert.IsType<QuoteBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesSupplementalFigureBlocksAroundDirectPicture() {
        const string html = """
<figure>
  <picture>
    <source srcset="/img/hero.webp 1x, /img/hero@2x.webp 2x" />
    <img alt="Hero" />
  </picture>
  <figcaption>Hero image</figcaption>
  <p>Photo credit: Team</p>
</figure>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        Assert.Collection(document.Blocks,
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/img/hero.webp", image.Path);
                Assert.Equal("Hero", image.Alt);
                Assert.Equal("Hero image", image.Caption);
            },
            block => Assert.Equal("Photo credit: Team", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("_Hero image_", markdown, StringComparison.Ordinal);
        Assert.Contains("Photo credit: Team", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesLinkedFigureMediaAsTypedImageBlock() {
        const string html = """
<figure>
  <a href="/docs/hero" title="Hero page" target="_blank" rel="nofollow sponsored">
    <img src="/img/hero.png" alt="Hero" title="View hero" />
  </a>
  <figcaption>Hero image</figcaption>
</figure>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/hero.png", image.Path);
        Assert.Equal("Hero", image.Alt);
        Assert.Equal("View hero", image.Title);
        Assert.Equal("https://example.com/docs/hero", image.LinkUrl);
        Assert.Equal("Hero page", image.LinkTitle);
        Assert.Equal("_blank", image.LinkTarget);
        Assert.Equal("nofollow sponsored", image.LinkRel);
        Assert.Equal("Hero image", image.Caption);

        string markdown = document.ToMarkdown();
        Assert.Contains("[![Hero](https://example.com/img/hero.png \"View hero\")](https://example.com/docs/hero \"Hero page\")", markdown, StringComparison.Ordinal);

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<a href=\"https://example.com/docs/hero\" title=\"Hero page\" target=\"_blank\" rel=\"", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("nofollow", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("sponsored", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("noopener", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("noreferrer", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesWrappedLinkedPictureFigureMediaAndSupplementalBlocks() {
        const string html = """
<figure>
  <div class="figure-media">
    <a href="/docs/hero">
      <picture>
        <source srcset="/img/hero.webp 1x, /img/hero@2x.webp 2x" />
        <img alt="Hero" />
      </picture>
    </a>
  </div>
  <figcaption>Hero image</figcaption>
  <p>Photo credit: Team</p>
</figure>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        Assert.Collection(document.Blocks,
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/img/hero.webp", image.Path);
                Assert.Equal("Hero", image.Alt);
                Assert.Equal("https://example.com/docs/hero", image.LinkUrl);
                Assert.Equal("Hero image", image.Caption);
            },
            block => Assert.Equal("Photo credit: Team", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("[![Hero](https://example.com/img/hero.webp)](https://example.com/docs/hero)", markdown, StringComparison.Ordinal);
        Assert.Contains("_Hero image_", markdown, StringComparison.Ordinal);
        Assert.Contains("Photo credit: Team", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesAnchorWrappedImageFigureMediaAndSupplementalBlocks() {
        const string html = """
<figure>
  <a href="/docs/hero">
    <span class="media-frame">
      <img src="/img/hero.png" alt="Hero" title="View hero" />
    </span>
  </a>
  <figcaption>Hero image</figcaption>
  <p>Photo credit: Team</p>
</figure>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        Assert.Collection(document.Blocks,
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/img/hero.png", image.Path);
                Assert.Equal("Hero", image.Alt);
                Assert.Equal("View hero", image.Title);
                Assert.Equal("https://example.com/docs/hero", image.LinkUrl);
                Assert.Equal("Hero image", image.Caption);
            },
            block => Assert.Equal("Photo credit: Team", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("[![Hero](https://example.com/img/hero.png \"View hero\")](https://example.com/docs/hero)", markdown, StringComparison.Ordinal);
        Assert.Contains("_Hero image_", markdown, StringComparison.Ordinal);
        Assert.Contains("Photo credit: Team", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesParagraphBreaksInsideTableCells() {
        string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><p>First</p><p>Second</p></td></tr></table>";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal("Alpha", table.Rows[0][0]);
        Assert.Equal("First\n\nSecond", table.Rows[0][1]);
        Assert.Collection(table.RowCells[0][1].Blocks,
            block => Assert.Equal("First", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Second", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("First<br><br>Second", markdown, StringComparison.Ordinal);

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<td><p>First</p><p>Second</p></td>", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesNestedTablesInsideOwningCellAst() {
        const string html = """
<table>
  <tr><th>Outer</th></tr>
  <tr>
    <td>
      <table>
        <tr><th>Inner</th></tr>
        <tr><td>Cell</td></tr>
      </table>
    </td>
  </tr>
</table>
""";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal(new[] { "Outer" }, table.Headers);
        Assert.Single(table.Rows);
        Assert.Single(table.RowCells);
        Assert.Single(table.RowCells[0]);

        var nestedTable = Assert.IsType<TableBlock>(Assert.Single(table.RowCells[0][0].Blocks));
        Assert.Equal(new[] { "Inner" }, nestedTable.Headers);
        Assert.Single(nestedTable.Rows);
        Assert.Equal("Cell", nestedTable.Rows[0][0]);
    }

    [Fact]
    public void HtmlToMarkdown_ReadsTableAlignmentFromInlineStyle() {
        const string html = """
<table>
  <tr>
    <th style="text-align: center;">Name</th>
    <th style="text-align:right">Count</th>
  </tr>
  <tr>
    <td>Alpha</td>
    <td>42</td>
  </tr>
</table>
""";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal(new[] { ColumnAlignment.Center, ColumnAlignment.Right }, table.Alignments);

        string markdown = document.ToMarkdown();
        Assert.Contains("| :---: | ---: |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_RecoversLazyLoadedImageSourcesAndStyleDimensions() {
        const string html = "<figure><img data-src=\"/img/demo.png\" alt=\"Demo\" style=\"width: 640px; height: 480px;\" /></figure>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/demo.png", image.Path);
        Assert.Equal("Demo", image.Alt);
        Assert.Equal(640d, image.Width);
        Assert.Equal(480d, image.Height);

        string markdown = document.ToMarkdown();
        Assert.Contains("width=640", markdown, StringComparison.Ordinal);
        Assert.Contains("height=480", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsPictureElementUsingSourceSetFallback() {
        const string html = """
<picture>
  <source srcset="/img/hero.webp 1x, /img/hero@2x.webp 2x" type="image/webp" />
  <img alt="Hero" />
</picture>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/hero.webp", image.Path);
        Assert.Equal("Hero", image.Alt);
    }

    [Fact]
    public void HtmlToMarkdown_CapturesFigurePictureCaptionAndLazySource() {
        const string html = """
<figure>
  <picture>
    <source data-srcset="/img/diagram.png 1x, /img/diagram@2x.png 2x" />
    <img alt="Diagram" />
  </picture>
  <figcaption>Architecture overview</figcaption>
</figure>
""";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/diagram.png", image.Path);
        Assert.Equal("Diagram", image.Alt);
        Assert.Equal("Architecture overview", image.Caption);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedBlockContentInsideTableCells() {
        string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><p>Intro</p><blockquote><p>Quoted</p></blockquote></td></tr></table>";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Collection(table.RowCells[0][1].Blocks,
            block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.IsType<QuoteBlock>(block));

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<td><p>Intro</p><blockquote><p>Quoted</p></blockquote></td>", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Recovers_Rendered_Callout_Block() {
        string html = "<blockquote class=\"callout note\"><p><strong>Important</strong></p><p>Body</p><ul><li>Nested</li></ul></blockquote>";

        MarkdownDoc document = html.LoadFromHtml();
        var callout = Assert.IsType<CalloutBlock>(Assert.Single(document.Blocks));

        Assert.Equal("note", callout.Kind);
        Assert.Equal("Important", callout.TitleInlines.RenderMarkdown());
        Assert.Collection(callout.ChildBlocks,
            block => Assert.Equal("Body", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.IsType<UnorderedListBlock>(block));

        string markdown = document.ToMarkdown();
        Assert.Contains("> [!NOTE] Important", markdown, StringComparison.Ordinal);
        Assert.Contains("> Body", markdown, StringComparison.Ordinal);
        Assert.Contains("> - Nested", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Roundtrips_Rendered_Callout_Block_Without_Synthesizing_Default_Title() {
        var document = MarkdownReader.Parse("""
> [!NOTE]
> Body
""");

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        MarkdownDoc roundtripped = renderedHtml.LoadFromHtml();

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(roundtripped.Blocks));
        Assert.Equal("note", callout.Kind);
        Assert.Equal(string.Empty, callout.TitleInlines.RenderMarkdown());
        Assert.Equal("Body", Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks)).Inlines.RenderMarkdown());

        string markdown = roundtripped.ToMarkdown();
        Assert.Contains("> [!NOTE]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("> [!NOTE] Note", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("> [!NOTE] note", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Recovers_Rendered_Callout_Block_Inside_Table_Cell() {
        string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><blockquote class=\"callout warning\"><p><strong>Watch</strong></p><p>Body</p></blockquote></td></tr></table>";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(table.RowCells[0][1].Blocks));
        Assert.Equal("warning", callout.Kind);
        Assert.Equal("Watch", callout.TitleInlines.RenderMarkdown());
        Assert.Equal("Body", Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks)).Inlines.RenderMarkdown());

        string markdown = document.ToMarkdown();
        Assert.Contains("[!WARNING] Watch", markdown, StringComparison.Ordinal);

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("class=\"callout warning\"", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Roundtrips_Rendered_Untitled_Callout_Block_Inside_Table_Cell() {
        var document = MarkdownReader.Parse("""
| Notes |
| --- |
| > [!WARNING]<br>> Body |
""");

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        MarkdownDoc roundtripped = renderedHtml.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(roundtripped.Blocks));

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(table.RowCells[0][0].Blocks));
        Assert.Equal("warning", callout.Kind);
        Assert.Equal(string.Empty, callout.TitleInlines.RenderMarkdown());
        Assert.Equal("Body", Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks)).Inlines.RenderMarkdown());

        string markdown = roundtripped.ToMarkdown();
        Assert.Contains("[!WARNING]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[!WARNING] Warning", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Converts_LinkWrapped_Lazy_Image_Paragraph_To_LinkedImageBlock() {
        string html = """
<p><a href="https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png"><img class="aligncenter size-full wp-image-4510 ewww_webp_lazy_load" src="data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20440%20482'%3E%3C/svg%3E" alt="GPO Registry Add" width="440" height="482" data-lazy-srcset="https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png 440w, https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd-274x300.png 274w" data-lazy-sizes="(max-width: 440px) 100vw, 440px" data-lazy-src="https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png" /><noscript><img class="aligncenter size-full wp-image-4510" src="https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png" alt="GPO Registry Add" width="440" height="482" srcset="https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png 440w, https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd-274x300.png 274w" sizes="(max-width: 440px) 100vw, 440px" /></noscript></a></p>
""";

        MarkdownDoc document = html.LoadFromHtml();
        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));

        Assert.Equal("https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png", image.Path);
        Assert.Equal("GPO Registry Add", image.Alt);
        Assert.Equal(440d, image.Width);
        Assert.Equal(482d, image.Height);
        Assert.Equal("https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png", image.LinkUrl);

        string markdown = document.ToMarkdown();
        Assert.Contains("[![GPO Registry Add](https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png)](https://example.com/wp-content/uploads/2015/08/GPO_RegistryAdd.png){width=440 height=482}", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/svg+xml", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("<noscript>", markdown, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToMarkdown_Converts_DivWrapped_Lazy_Linked_Image_To_ImageBlock_Instead_Of_RawHtml() {
        string html = """
<div class="post-img"><a href="https://example.com/automating-network-diagnostics-with-globalping-powershell-module/" class="default"><img width="256" height="256" src="data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20256%20256'%3E%3C/svg%3E" class="img-responsive wp-post-image" alt="Automating Network Diagnostics with Globalping PowerShell Module" data-lazy-srcset="https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-small.jpg 1x, https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-large.jpg 2x" data-lazy-sizes="100vw" data-lazy-src="https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-large.jpg" /><noscript><img width="256" height="256" src="https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-large.jpg" class="img-responsive wp-post-image" alt="Automating Network Diagnostics with Globalping PowerShell Module" srcset="https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-small.jpg 1x, https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-large.jpg 2x" sizes="100vw" /></noscript></a></div>
""";

        MarkdownDoc document = html.LoadFromHtml();
        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));

        Assert.Equal("https://example.com/wp-content/uploads/2025/06/Automating-Network-Diagnostics-with-Globalping-PowerShell-Module-thegem-post-thumb-large.jpg", image.Path);
        Assert.Equal("Automating Network Diagnostics with Globalping PowerShell Module", image.Alt);
        Assert.Equal(256d, image.Width);
        Assert.Equal(256d, image.Height);
        Assert.Equal("https://example.com/automating-network-diagnostics-with-globalping-powershell-module/", image.LinkUrl);

        string markdown = document.ToMarkdown();
        Assert.Contains("Automating Network Diagnostics with Globalping PowerShell Module", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("<a href=", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<noscript>", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("data:image/svg+xml", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_DoesNotInvent_RawUrl_For_Empty_Overlay_Anchor() {
        string html = """
<div class="related-element">
  <a href="https://example.com/exchange-2013-integration-with-sharepoint-doesnt-work/" aria-label="Exchange 2013 integration with SharePoint doesn’t work"><span class="gem-dummy"></span></a>
  <div class="related-element-info clearfix">
    <div class="related-element-info-conteiner">
      <a href="https://example.com/exchange-2013-integration-with-sharepoint-doesnt-work/">Exchange 2013 integration with SharePoint doesn’t work</a>
      <div class="related-element-info-excerpt">
        <p>The steps to integrate new Microsoft Exchange 2013 with SharePoint 2013 are fairly simple.</p>
      </div>
    </div>
    <div class="post-meta date-color">
      <div class="entry-meta clearfix">
        <div class="post-meta-right">
          <span class="comments-link"><a href="https://example.com/exchange-2013-integration-with-sharepoint-doesnt-work/#respond">0</a></span>
        </div>
        <div class="post-meta-left">
          <span class="post-meta-date gem-post-date gem-date-color small-body">14 Jun 2015</span>
        </div>
      </div>
    </div>
  </div>
</div>
""";

        MarkdownDoc document = html.LoadFromHtml();
        string markdown = document.ToMarkdown();

        Assert.DoesNotContain("[https://example.com/exchange-2013-integration-with-sharepoint-doesnt-work/]", markdown, StringComparison.Ordinal);
        Assert.Contains("[Exchange 2013 integration with SharePoint doesn’t work](https://example.com/exchange-2013-integration-with-sharepoint-doesnt-work/)", markdown, StringComparison.Ordinal);
        Assert.Contains("The steps to integrate new Microsoft Exchange 2013 with SharePoint 2013 are fairly simple.", markdown, StringComparison.Ordinal);
        Assert.Contains("[0](https://example.com/exchange-2013-integration-with-sharepoint-doesnt-work/#respond)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Preserves_Repeated_Listing_Card_Metadata_By_Default() {
        string html = """
<main>
  <div class="blog-feed">
    <article class="post-card">
      <div class="post-date">15 June</div>
      <div class="post-time">21:52</div>
      <div class="entry-meta"><span class="post-meta-author">By Przemyslaw Klys</span></div>
      <div class="post-title"><h3><a href="https://example.com/post-one" rel="bookmark">15 Jun: First post</a></h3></div>
      <div class="summary"><p>First summary.</p></div>
      <div class="post-read-more"><a href="https://example.com/post-one">Read More</a></div>
    </article>
    <article class="post-card">
      <div class="post-date">14 June</div>
      <div class="post-time">19:20</div>
      <div class="entry-meta"><span class="post-meta-author">By Another Author</span></div>
      <div class="post-title"><h3><a href="https://example.com/post-two" rel="bookmark">14 Jun: Second post</a></h3></div>
      <div class="summary"><p>Second summary.</p></div>
      <div class="post-read-more"><a href="https://example.com/post-two">Read More</a></div>
    </article>
  </div>
</main>
""";

        string markdown = html.ToMarkdown();

        Assert.Contains("15 June", markdown, StringComparison.Ordinal);
        Assert.Contains("21:52", markdown, StringComparison.Ordinal);
        Assert.Contains("By Przemyslaw Klys", markdown, StringComparison.Ordinal);
        Assert.Contains("[Read More](https://example.com/post-one)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_CanSuppress_Repeated_Listing_Card_Metadata() {
        string html = """
<main>
  <div class="blog-feed">
    <article class="post-card">
      <div class="post-date">15 June</div>
      <div class="post-time">21:52</div>
      <div class="entry-meta"><span class="post-meta-author">By Przemyslaw Klys</span></div>
      <div class="post-title"><h3><a href="https://example.com/post-one" rel="bookmark">15 Jun: First post</a></h3></div>
      <div class="summary"><p>First summary.</p></div>
      <div class="post-read-more"><a href="https://example.com/post-one">Read More</a></div>
    </article>
    <article class="post-card">
      <div class="post-date">14 June</div>
      <div class="post-time">19:20</div>
      <div class="entry-meta"><span class="post-meta-author">By Another Author</span></div>
      <div class="post-title"><h3><a href="https://example.com/post-two" rel="bookmark">14 Jun: Second post</a></h3></div>
      <div class="summary"><p>Second summary.</p></div>
      <div class="post-read-more"><a href="https://example.com/post-two">Read More</a></div>
    </article>
  </div>
</main>
""";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            ListingCardMetadataMode = HtmlListingCardMetadataMode.SuppressInRepeatedCards
        });

        Assert.DoesNotContain("15 June", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("21:52", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("By Przemyslaw Klys", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[Read More](https://example.com/post-one)", markdown, StringComparison.Ordinal);
        Assert.Contains("[15 Jun: First post](https://example.com/post-one)", markdown, StringComparison.Ordinal);
        Assert.Contains("First summary.", markdown, StringComparison.Ordinal);
        Assert.Contains("[14 Jun: Second post](https://example.com/post-two)", markdown, StringComparison.Ordinal);
        Assert.Contains("Second summary.", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedListItemBlockOrder() {
        string html = "<ul><li><p>Alpha</p><blockquote><p>Quoted</p></blockquote><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int quoteIndex = markdown.IndexOf("Quoted", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(quoteIndex > alphaIndex, "Expected quoted content after the opening paragraph.");
        Assert.True(omegaIndex > quoteIndex, "Expected trailing paragraph after the quote block.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleDefinitionsPerTerm() {
        string html = "<dl><dt>Term</dt><dd>First definition</dd><dd>Second definition</dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        var group = Assert.Single(list.Groups);
        Assert.Single(group.Terms);
        Assert.Equal(2, group.Definitions.Count);
        Assert.Equal(2, list.Entries.Count);
        Assert.Equal("Term", list.Entries[0].Term.RenderMarkdown());
        Assert.Equal("First definition", Assert.IsType<ParagraphBlock>(Assert.Single(list.Entries[0].DefinitionBlocks)).Inlines.RenderMarkdown());
        Assert.Equal("Term", list.Entries[1].Term.RenderMarkdown());
        Assert.Equal("Second definition", Assert.IsType<ParagraphBlock>(Assert.Single(list.Entries[1].DefinitionBlocks)).Inlines.RenderMarkdown());
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleParagraphsInDefinitionValues() {
        string html = "<dl><dt>Term</dt><dd><p>First paragraph</p><p>Second paragraph</p></dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Single(list.Entries);
        Assert.Equal("Term", list.Entries[0].Term.RenderMarkdown());
        Assert.Collection(list.Entries[0].DefinitionBlocks,
            block => Assert.Equal("First paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Second paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("Term: First paragraph", markdown, StringComparison.Ordinal);
        Assert.Contains("Second paragraph", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedBlockContentInDefinitionValues() {
        string html = "<dl><dt>Term</dt><dd><p>Intro</p><blockquote><p>Quoted</p></blockquote><ul><li>Nested</li></ul></dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Single(list.Entries);
        Assert.Equal("Term", list.Entries[0].Term.RenderMarkdown());
        Assert.Collection(list.Entries[0].DefinitionBlocks,
            block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.IsType<UnorderedListBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesNestedListOrderInsideListItem() {
        string html = "<ul><li><p>Alpha</p><ul><li>Nested</li></ul><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<UnorderedListBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int nestedIndex = markdown.IndexOf("Nested", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(nestedIndex > alphaIndex, "Expected nested list content after the opening paragraph.");
        Assert.True(omegaIndex > nestedIndex, "Expected trailing paragraph after the nested list.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesDetailsOrderInsideListItem() {
        string html = "<ul><li><p>Alpha</p><details open><summary>More</summary><p>Hidden</p></details><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<DetailsBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleTermsPerDefinitionGroup() {
        string html = "<dl><dt>Alpha</dt><dt>Beta</dt><dd>Shared definition</dd><dd>Follow-up definition</dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        var group = Assert.Single(list.Groups);
        Assert.Equal(new[] { "Alpha", "Beta" }, group.Terms.Select(term => term.RenderMarkdown()).ToArray());
        Assert.Equal(new[] { "Shared definition", "Follow-up definition" }, group.Definitions.Select(definition => definition.Markdown).ToArray());
        Assert.Equal(4, list.Entries.Count);
        Assert.Equal(("Alpha", "Shared definition"), list.Items[0]);
        Assert.Equal(("Beta", "Shared definition"), list.Items[1]);
        Assert.Equal(("Alpha", "Follow-up definition"), list.Items[2]);
        Assert.Equal(("Beta", "Follow-up definition"), list.Items[3]);
        Assert.Equal("Alpha", list.Entries[0].Term.RenderMarkdown());
        Assert.Equal("Shared definition", Assert.IsType<ParagraphBlock>(Assert.Single(list.Entries[0].DefinitionBlocks)).Inlines.RenderMarkdown());

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        int alphaIndex = renderedHtml.IndexOf("<dt>Alpha</dt>", StringComparison.Ordinal);
        int betaIndex = renderedHtml.IndexOf("<dt>Beta</dt>", StringComparison.Ordinal);
        int sharedIndex = renderedHtml.IndexOf("<dd>Shared definition</dd>", StringComparison.Ordinal);
        int followUpIndex = renderedHtml.IndexOf("<dd>Follow-up definition</dd>", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha term in rendered HTML.");
        Assert.True(betaIndex > alphaIndex, "Expected grouped term order before definitions.");
        Assert.True(sharedIndex > betaIndex, "Expected shared definition after grouped terms.");
        Assert.True(followUpIndex > sharedIndex, "Expected follow-up definition after shared definition.");
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsSharedVisualHostIntoSemanticFencedBlock() {
        var payload = MarkdownVisualContract.CreatePayload("{\"type\":\"bar\"}");
        string html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            MarkdownSemanticKinds.Chart,
            "vendor-chart",
            payload);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("vendor-chart", block.Language);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
        Assert.Equal(
            NormalizeMarkdown("```vendor-chart\n{\"type\":\"bar\"}\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsSharedVisualHostIntoSemanticFencedBlock_With_Fence_Metadata() {
        var payload = MarkdownVisualContract.CreatePayload("{\"type\":\"bar\"}");
        var fenceInfo = MarkdownCodeFenceInfo.Parse("vendor-chart #quarterly-summary .wide .accent title=\"Quarterly Overview\" pinned");
        string html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            MarkdownSemanticKinds.Chart,
            "vendor-chart",
            payload,
            fenceInfo);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("vendor-chart", block.Language);
        Assert.Equal("vendor-chart #quarterly-summary .wide .accent title=\"Quarterly Overview\" pinned", block.InfoString);
        Assert.Equal("quarterly-summary", block.FenceInfo.ElementId);
        Assert.Equal(new[] { "wide", "accent" }, block.FenceInfo.Classes);
        Assert.Equal("Quarterly Overview", block.FenceInfo.Title);
        Assert.Equal("true", block.FenceInfo.Attributes["pinned"]);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
        Assert.Equal(
            NormalizeMarkdown("```vendor-chart #quarterly-summary .wide .accent title=\"Quarterly Overview\" pinned\n{\"type\":\"bar\"}\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_Preserves_Figure_Caption_On_Shared_Visual_Hosts() {
        var payload = MarkdownVisualContract.CreatePayload("{\"type\":\"bar\"}");
        var host = MarkdownVisualContract.BuildElementHtml(
            "figure",
            "omd-visual omd-custom",
            MarkdownSemanticKinds.Chart,
            "vendor-chart",
            payload);
        string html = host.Replace("</figure>", "<figcaption>Quarterly Overview</figcaption></figure>");

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("vendor-chart", block.Language);
        Assert.Equal("Quarterly Overview", block.Caption);
        Assert.Equal(
            NormalizeMarkdown("```vendor-chart\n{\"type\":\"bar\"}\n```\n_Quarterly Overview_"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_Can_Use_Plugin_Provided_Visual_RoundTrip_Hints() {
        var plugin = new MarkdownRendererPlugin(
            "Vendor Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            visualElementRoundTripHints: new[] {
                new MarkdownVisualElementRoundTripHint(
                    "vendor.caption",
                    "Vendor caption",
                    context => context.VisualElement.TryGetAttribute("data-vendor-caption", out var caption)
                        && !string.IsNullOrWhiteSpace(caption)
                        ? context.CreateBlock(caption: caption)
                        : null)
            });
        var options = new HtmlToMarkdownOptions();
        options.ApplyPlugin(plugin);
        var payload = MarkdownVisualContract.CreatePayload("{\"type\":\"bar\"}");
        string html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            MarkdownSemanticKinds.Chart,
            "vendor-chart",
            payload,
            new KeyValuePair<string, string?>("data-vendor-caption", "Quarterly Overview"));

        MarkdownDoc document = html.LoadFromHtml(options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal("Quarterly Overview", block.Caption);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_Sample_StatusPanel_Package_EndToEnd() {
        const string raw = """
{"title":"Operations Overview","summary":"All checks passing","status":"healthy","caption":"Panel caption"}
""";
        var renderOptions = MarkdownRendererPresets.CreateStrictMinimal();
        renderOptions.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);
        var htmlToMarkdownOptions = new HtmlToMarkdownOptions();
        htmlToMarkdownOptions.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        string html = MarkdownRendererShell.RenderBodyHtml("```status-panel\n" + raw + "\n```", renderOptions);
        MarkdownDoc document = html.LoadFromHtml(htmlToMarkdownOptions);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal("status-panel", block.SemanticKind);
        Assert.Equal("status-panel", block.Language);
        Assert.Equal("status-panel title=\"Operations Overview\"", block.InfoString);
        Assert.Equal("Panel caption", block.Caption);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```status-panel title=\"Operations Overview\"\n" + raw + "\n```\n_Panel caption_"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_Applies_Sample_StatusPanel_Document_Transform_To_Standard_Code_Html() {
        const string html = """
<pre><code class="language-status-panel">{"title":"Operations Overview","summary":"All checks passing"}</code></pre>
""";
        var options = new HtmlToMarkdownOptions();
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        MarkdownDoc document = html.LoadFromHtml(options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal("status-panel", block.SemanticKind);
        Assert.Equal("status-panel", block.Language);
        Assert.Equal("{\"title\":\"Operations Overview\",\"summary\":\"All checks passing\"}", block.Content);
        Assert.Equal(
            NormalizeMarkdown("```status-panel\n{\"title\":\"Operations Overview\",\"summary\":\"All checks passing\"}\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_Applies_Sample_StatusPanel_ElementBlock_Converter_To_Vendor_Html() {
        const string payload = "{\"title\":\"Operations Overview\",\"summary\":\"All checks passing\",\"status\":\"healthy\"}";
        string html = """
<section class="sample-status-panel" data-sample-status-panel-json="{&quot;title&quot;:&quot;Operations Overview&quot;,&quot;summary&quot;:&quot;All checks passing&quot;,&quot;status&quot;:&quot;healthy&quot;}">
  <header>Operations Overview</header>
  <div class="sample-status-panel-body">
    <p>All checks passing</p>
  </div>
  <footer>Panel caption</footer>
</section>
""";
        var options = new HtmlToMarkdownOptions();
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        MarkdownDoc document = html.LoadFromHtml(options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal("status-panel", block.SemanticKind);
        Assert.Equal("status-panel", block.Language);
        Assert.Equal("status-panel title=\"Operations Overview\"", block.InfoString);
        Assert.Equal(payload, block.Content);
        Assert.Equal("Panel caption", block.Caption);
        Assert.Equal(
            NormalizeMarkdown("```status-panel title=\"Operations Overview\"\n" + payload + "\n```\n_Panel caption_"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_Preserves_Opaque_Fence_Metadata_Tail_From_Shared_Visual_Host() {
        var payload = MarkdownVisualContract.CreatePayload("{\"type\":\"bar\"}");
        var fenceInfo = MarkdownCodeFenceInfo.Parse("vendor-chart {#quarterly-summary .wide title=\"Quarterly Overview\"");
        string html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            MarkdownSemanticKinds.Chart,
            "vendor-chart",
            payload,
            fenceInfo);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("vendor-chart", block.Language);
        Assert.Equal("vendor-chart {#quarterly-summary .wide title=\"Quarterly Overview\"", block.InfoString);
        Assert.Null(block.FenceInfo.ElementId);
        Assert.Empty(block.FenceInfo.Classes);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
        Assert.Equal(
            NormalizeMarkdown("```vendor-chart {#quarterly-summary .wide title=\"Quarterly Overview\"\n{\"type\":\"bar\"}\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedIxDataviewHtml_BackToSemanticFence() {
        const string raw = """
{"title":"Replication Summary","summary":"Latest replication posture","kind":"ix_tool_dataview_v1","call_id":"call_123","headers":["Domain","Status"],"items":[{"Domain":"ad.evotec.xyz","Status":"Healthy"}]}
""";
        var options = MarkdownRendererPresets.CreateStrictMinimal();
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        MarkdownRendererIntelligenceXLegacyMigration.Apply(options);
        string html = MarkdownRendererShell.RenderBodyHtml("```ix-dataview\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.DataView, block.SemanticKind);
        Assert.Equal("ix-dataview", block.Language);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```ix-dataview\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedIxNetworkHtml_BackToSemanticFence() {
        const string raw = """
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Network.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```ix-network\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Network, block.SemanticKind);
        Assert.Equal("ix-network", block.Language);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```ix-network\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedIxNetworkHtml_BackToSemanticFence_With_Fence_Metadata() {
        const string raw = """
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Network.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```ix-network #relationship-map .wide title=\"Relationship Map\" pinned\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Network, block.SemanticKind);
        Assert.Equal("ix-network", block.Language);
        Assert.Equal("ix-network #relationship-map .wide title=\"Relationship Map\" pinned", block.InfoString);
        Assert.Equal("relationship-map", block.FenceInfo.ElementId);
        Assert.Equal(new[] { "wide" }, block.FenceInfo.Classes);
        Assert.Equal("Relationship Map", block.FenceInfo.Title);
        Assert.Equal("true", block.FenceInfo.Attributes["pinned"]);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```ix-network #relationship-map .wide title=\"Relationship Map\" pinned\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedMermaidHtml_BackToSemanticFence() {
        const string raw = """
flowchart LR
A[Markdown Input] --> B{Parser OK?}
B --> C[Render Mermaid]
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Mermaid.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```mermaid\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Mermaid, block.SemanticKind);
        Assert.Equal("mermaid", block.Language);
        Assert.Equal(NormalizeMarkdown(raw), NormalizeMarkdown(block.Content));
        Assert.Equal(
            NormalizeMarkdown("```mermaid\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedMathHtml_BackToSemanticFence() {
        const string raw = """
\sum_{i=1}^{n} x_i^2
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Math.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```math\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Math, block.SemanticKind);
        Assert.Equal("math", block.Language);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```math\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_MixedTranscriptFragment_WithVisuals_AndStructuredBlocks() {
        const string markdown = """
# Transcript Sample

## Proactive checks

| Check | Expected |
|---|---|
| Fence integrity | Visuals render |

```ix-chart
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
```

```ix-network
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
```

```mermaid
flowchart LR
A --> B
```

```math
x^2 + 1
```

1. Item one
2. Item two
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        options.Mermaid.Enabled = true;
        options.Math.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.Collection(document.Blocks,
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<TableBlock>(block),
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("ix-chart", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Chart, semantic.SemanticKind);
            },
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("ix-network", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Network, semantic.SemanticKind);
            },
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("mermaid", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Mermaid, semantic.SemanticKind);
            },
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("math", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Math, semantic.SemanticKind);
            },
            block => Assert.IsType<OrderedListBlock>(block));

        string roundTripped = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("```ix-chart", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```ix-network", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```mermaid", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```math", roundTripped, StringComparison.Ordinal);
        Assert.Contains("| Check | Expected |", roundTripped, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_ExportedTranscriptVisualPack_WithSemanticRecovery() {
        string markdown = LoadCompatibilityFixture("ix-exported-transcript-visual-pack.md");
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        options.Mermaid.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.True(document.Blocks.OfType<HeadingBlock>().Count() >= 8);
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:30: 13)");
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "User (20:36: 21)");
        Assert.True(document.Blocks.OfType<TableBlock>().Count() >= 1);
        Assert.True(document.Blocks.OfType<OrderedListBlock>().Count() >= 1);

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(5, semanticBlocks.Count(block => block.Language == "ix-chart"));
        Assert.Equal(2, semanticBlocks.Count(block => block.Language == "mermaid"));
        Assert.Contains(semanticBlocks, block => block.SemanticKind == MarkdownSemanticKinds.Chart && block.Language == "ix-chart");
        Assert.Contains(semanticBlocks, block => block.SemanticKind == MarkdownSemanticKinds.Mermaid && block.Language == "mermaid");
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Broken\"", StringComparison.Ordinal));

        string roundTripped = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("## Proactive checks", roundTripped, StringComparison.Ordinal);
        Assert.Contains("| Check | What to verify | Expected pass signal |", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```ix-chart", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```mermaid", roundTripped, StringComparison.Ordinal);
        Assert.Contains("\"label\": \"Broken\"", roundTripped, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_ExportedTranscriptVisualPack_PreservesSemanticBlockOrder() {
        string markdown = LoadCompatibilityFixture("ix-exported-transcript-visual-pack.md");
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        options.Mermaid.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        var sequence = document.Blocks
            .Where(block => block is SemanticFencedBlock or TableBlock or OrderedListBlock)
            .Select(static block => block switch {
                SemanticFencedBlock semantic => $"semantic:{semantic.Language}",
                TableBlock => "table",
                OrderedListBlock => "olist",
                _ => "other"
            })
            .ToArray();

        Assert.Equal(
            [
                "table",
                "olist",
                "semantic:ix-chart",
                "semantic:mermaid",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:mermaid"
            ],
            sequence);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_ExportedTranscriptChartSuite_WithSemanticRecovery() {
        string markdown = LoadCompatibilityFixture("ix-exported-transcript-chart-suite.md");
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Mermaid.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:36: 24)");
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:37: 49)");
        Assert.True(document.Blocks.OfType<TableBlock>().Count() >= 1);

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(4, semanticBlocks.Count(block => block.Language == "ix-chart"));
        Assert.Equal(1, semanticBlocks.Count(block => block.Language == "mermaid"));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Critical\"", StringComparison.Ordinal));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Privileged Changes\"", StringComparison.Ordinal));

        var sequence = document.Blocks
            .Where(block => block is SemanticFencedBlock or TableBlock)
            .Select(static block => block switch {
                SemanticFencedBlock semantic => $"semantic:{semantic.Language}",
                TableBlock => "table",
                _ => "other"
            })
            .ToArray();

        Assert.Equal(
            [
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "table",
                "semantic:mermaid"
            ],
            sequence);

        string roundTripped = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("Risk distribution", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```ix-chart", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```mermaid", roundTripped, StringComparison.Ordinal);
        Assert.Contains("Feature | Test block | Expected result | If it fails, likely cause", roundTripped, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithResolvedLinkedResponsiveFigure() {
        string html = LoadHtmlFixture("publisher-linked-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Storm Update", Assert.IsType<HeadingBlock>(block).Text),
            block => {
                string markdown = Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown();
                Assert.Contains("[briefing](https://example.com/news/2026/briefing.html)", markdown, StringComparison.Ordinal);
            },
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/news/2026/media/storm-center.webp", image.Path);
                Assert.Equal("Flooded street at dawn", image.Alt);
                Assert.Equal("Open full photo", image.Title);
                Assert.Equal(1280d, image.Width);
                Assert.Equal(720d, image.Height);
                Assert.Equal("https://example.com/news/2026/gallery/storm-center", image.LinkUrl);
                Assert.Equal("Open photo", image.LinkTitle);
                Assert.Equal("_blank", image.LinkTarget);
                Assert.Equal("nofollow sponsored", image.LinkRel);
                Assert.Equal("Residents navigate floodwater after the overnight storm.", image.Caption);
            },
            block => Assert.Equal("Photo credit: City Desk", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => {
                string markdown = Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown();
                Assert.Contains("[flood map](https://example.com/news/maps/flood-zones.html)", markdown, StringComparison.Ordinal);
            });

        string renderedMarkdown = document.ToMarkdown();
        Assert.Contains("[![Flooded street at dawn](https://example.com/news/2026/media/storm-center.webp \"Open full photo\")](https://example.com/news/2026/gallery/storm-center \"Open photo\")", renderedMarkdown, StringComparison.Ordinal);
        Assert.Contains("_Residents navigate floodwater after the overnight storm._", renderedMarkdown, StringComparison.Ordinal);
        Assert.Contains("Photo credit: City Desk", renderedMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithResolvedLinkedResponsiveFigure_EmitsStableMarkdownSnapshot() {
        string html = LoadHtmlFixture("publisher-linked-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        const string expected = """
# Storm Update

Read the [briefing](https://example.com/news/2026/briefing.html) before crews enter the river district.

[![Flooded street at dawn](https://example.com/news/2026/media/storm-center.webp "Open full photo")](https://example.com/news/2026/gallery/storm-center "Open photo"){width=1280 height=720}
_Residents navigate floodwater after the overnight storm._

Photo credit: City Desk

Inspect the [flood map](https://example.com/news/maps/flood-zones.html) for the latest street closures.
""";

        Assert.Equal(NormalizeMarkdown(expected), NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithNoscriptResponsiveFallback() {
        string html = LoadHtmlFixture("publisher-noscript-linked-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Storm Update", Assert.IsType<HeadingBlock>(block).Text),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/news/2026/media/storm-center.webp", image.Path);
                Assert.Equal("Flooded street at dawn", image.Alt);
                Assert.Equal("Open full photo", image.Title);
                Assert.Equal(1280d, image.Width);
                Assert.Equal(720d, image.Height);
                Assert.Equal("https://example.com/news/2026/gallery/storm-center", image.LinkUrl);
                Assert.Equal("Open photo", image.LinkTitle);
                Assert.Equal("_blank", image.LinkTarget);
                Assert.Equal("nofollow sponsored", image.LinkRel);
                Assert.Equal("Residents navigate floodwater after the overnight storm.", image.Caption);
            },
            block => Assert.Equal("Photo credit: City Desk", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string renderedMarkdown = document.ToMarkdown();
        Assert.DoesNotContain("data:image/gif", renderedMarkdown, StringComparison.Ordinal);
        Assert.Contains("storm-center.webp", renderedMarkdown, StringComparison.Ordinal);
        Assert.Contains("_Residents navigate floodwater after the overnight storm._", renderedMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithNoscriptResponsiveFallback_EmitsStableMarkdownSnapshot() {
        string html = LoadHtmlFixture("publisher-noscript-linked-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        const string expected = """
# Storm Update

[![Flooded street at dawn](https://example.com/news/2026/media/storm-center.webp "Open full photo")](https://example.com/news/2026/gallery/storm-center "Open photo"){width=1280 height=720}
_Residents navigate floodwater after the overnight storm._

Photo credit: City Desk
""";

        Assert.Equal(NormalizeMarkdown(expected), NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithArtDirectedPictureSources() {
        string html = LoadHtmlFixture("publisher-art-direction-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Storm Update", Assert.IsType<HeadingBlock>(block).Text),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/news/2026/media/storm-wide.webp", image.Path);
                Assert.Equal("https://example.com/news/2026/media/storm-fallback.jpg", image.PictureFallbackPath);
                Assert.Equal("Flooded street at dawn", image.Alt);
                Assert.Equal("Open full photo", image.Title);
                Assert.Equal(1280d, image.Width);
                Assert.Equal(720d, image.Height);
                Assert.Equal("Residents navigate floodwater after the overnight storm.", image.Caption);
                Assert.Collection(image.PictureSources,
                    source => {
                        Assert.Equal("https://example.com/news/2026/media/storm-wide.webp", source.Path);
                        Assert.Equal("https://example.com/news/2026/media/storm-wide.webp 1x, https://example.com/news/2026/media/storm-wide@2x.webp 2x", source.SrcSet);
                        Assert.Equal("(min-width: 960px)", source.Media);
                        Assert.Equal("image/webp", source.Type);
                        Assert.Equal("100vw", source.Sizes);
                    },
                    source => {
                        Assert.Equal("https://example.com/news/2026/media/storm-mobile.webp", source.Path);
                        Assert.Equal("https://example.com/news/2026/media/storm-mobile.webp 1x, https://example.com/news/2026/media/storm-mobile@2x.webp 2x", source.SrcSet);
                        Assert.Equal("(max-width: 959px)", source.Media);
                        Assert.Equal("image/webp", source.Type);
                        Assert.Equal("100vw", source.Sizes);
                    });
            });

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<picture>", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<source srcset=\"https://example.com/news/2026/media/storm-wide.webp 1x, https://example.com/news/2026/media/storm-wide@2x.webp 2x\" media=\"(min-width: 960px)\" type=\"image/webp\" sizes=\"100vw\" />", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<source srcset=\"https://example.com/news/2026/media/storm-mobile.webp 1x, https://example.com/news/2026/media/storm-mobile@2x.webp 2x\" media=\"(max-width: 959px)\" type=\"image/webp\" sizes=\"100vw\" />", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<img src=\"https://example.com/news/2026/media/storm-fallback.jpg\" alt=\"Flooded street at dawn\" title=\"Open full photo\" width=\"1280\" height=\"720\"", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithArtDirectedPictureSources_EmitsStableMarkdownSnapshot() {
        string html = LoadHtmlFixture("publisher-art-direction-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        const string expected = """
# Storm Update

![Flooded street at dawn](https://example.com/news/2026/media/storm-wide.webp "Open full photo"){width=1280 height=720}
_Residents navigate floodwater after the overnight storm._
""";

        Assert.Equal(NormalizeMarkdown(expected), NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithCdnLazyPictureSources() {
        string html = LoadHtmlFixture("publisher-cdn-lazy-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Storm Update", Assert.IsType<HeadingBlock>(block).Text),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://cdn.example.net/images/storm-wide.avif", image.Path);
                Assert.Equal("https://example.com/news/2026/media/storm-fallback.jpg", image.PictureFallbackPath);
                Assert.Equal("Flooded street at dawn", image.Alt);
                Assert.Equal("Open full photo", image.Title);
                Assert.Equal(1280d, image.Width);
                Assert.Equal(720d, image.Height);
                Assert.Equal("Residents navigate floodwater after the overnight storm.", image.Caption);
                Assert.Collection(image.PictureSources,
                    source => {
                        Assert.Equal("https://cdn.example.net/images/storm-wide.avif", source.Path);
                        Assert.Equal("https://cdn.example.net/images/storm-wide.avif 1x, https://cdn.example.net/images/storm-wide@2x.avif 2x", source.SrcSet);
                        Assert.Equal("(min-width: 960px)", source.Media);
                        Assert.Equal("image/avif", source.Type);
                        Assert.Equal("100vw", source.Sizes);
                    },
                    source => {
                        Assert.Equal("https://example.com/news/2026/media/storm-mobile.webp", source.Path);
                        Assert.Equal("https://example.com/news/2026/media/storm-mobile.webp 1x, https://example.com/news/2026/media/storm-mobile@2x.webp 2x", source.SrcSet);
                        Assert.Equal("(max-width: 959px)", source.Media);
                        Assert.Equal("image/webp", source.Type);
                        Assert.Equal("100vw", source.Sizes);
                    });
            });

        string renderedMarkdown = document.ToMarkdown();
        Assert.DoesNotContain("data:image/gif", renderedMarkdown, StringComparison.Ordinal);
        Assert.Contains("https://cdn.example.net/images/storm-wide.avif", renderedMarkdown, StringComparison.Ordinal);

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<source srcset=\"https://cdn.example.net/images/storm-wide.avif 1x, https://cdn.example.net/images/storm-wide@2x.avif 2x\" media=\"(min-width: 960px)\" type=\"image/avif\" sizes=\"100vw\" />", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<source srcset=\"https://example.com/news/2026/media/storm-mobile.webp 1x, https://example.com/news/2026/media/storm-mobile@2x.webp 2x\" media=\"(max-width: 959px)\" type=\"image/webp\" sizes=\"100vw\" />", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<img src=\"https://example.com/news/2026/media/storm-fallback.jpg\" alt=\"Flooded street at dawn\" title=\"Open full photo\" width=\"1280\" height=\"720\"", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithWidthDescriptorPictureSources() {
        string html = LoadHtmlFixture("publisher-width-descriptor-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Storm Update", Assert.IsType<HeadingBlock>(block).Text),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/news/2026/media/storm-wide.webp?fit=cover&crop=10,20,300,400", image.Path);
                Assert.Equal("https://example.com/news/2026/media/storm-fallback.jpg?download=1", image.PictureFallbackPath);
                Assert.Equal("Flooded street at dawn", image.Alt);
                Assert.Equal("Open full photo", image.Title);
                Assert.Equal(1280d, image.Width);
                Assert.Equal(720d, image.Height);
                Assert.Equal("Residents navigate floodwater after the overnight storm.", image.Caption);
                Assert.Collection(image.PictureSources,
                    source => {
                        Assert.Equal("https://example.com/news/2026/media/storm-wide.webp?fit=cover&crop=10,20,300,400", source.Path);
                        Assert.Equal("https://example.com/news/2026/media/storm-wide.webp?fit=cover&crop=10,20,300,400 640w, https://example.com/news/2026/media/storm-wide.webp?fit=cover&crop=20,40,600,800 1280w", source.SrcSet);
                        Assert.Equal("(min-width: 960px)", source.Media);
                        Assert.Equal("image/webp", source.Type);
                        Assert.Equal("(min-width: 960px) 90vw, 100vw", source.Sizes);
                    },
                    source => {
                        Assert.Equal("https://example.com/news/2026/media/storm-mobile.webp?fit=cover&crop=5,10,200,250", source.Path);
                        Assert.Equal("https://example.com/news/2026/media/storm-mobile.webp?fit=cover&crop=5,10,200,250 320w, https://example.com/news/2026/media/storm-mobile.webp?fit=cover&crop=10,20,400,500 640w", source.SrcSet);
                        Assert.Equal("(max-width: 959px)", source.Media);
                        Assert.Equal("image/webp", source.Type);
                        Assert.Equal("100vw", source.Sizes);
                    });
            });

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<source srcset=\"https://example.com/news/2026/media/storm-wide.webp?fit=cover&amp;crop=10,20,300,400 640w, https://example.com/news/2026/media/storm-wide.webp?fit=cover&amp;crop=20,40,600,800 1280w\" media=\"(min-width: 960px)\" type=\"image/webp\" sizes=\"(min-width: 960px) 90vw, 100vw\" />", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("<source srcset=\"https://example.com/news/2026/media/storm-mobile.webp?fit=cover&amp;crop=5,10,200,250 320w, https://example.com/news/2026/media/storm-mobile.webp?fit=cover&amp;crop=10,20,400,500 640w\" media=\"(max-width: 959px)\" type=\"image/webp\" sizes=\"100vw\" />", renderedHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("%20", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadsPublisherFixtureWithWidthDescriptorPictureSources_EmitsStableMarkdownSnapshot() {
        string html = LoadHtmlFixture("publisher-width-descriptor-picture-article.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        const string expected = """
# Storm Update

![Flooded street at dawn](https://example.com/news/2026/media/storm-wide.webp?fit=cover&crop=10,20,300,400 "Open full photo"){width=1280 height=720}
_Residents navigate floodwater after the overnight storm._
""";

        Assert.Equal(NormalizeMarkdown(expected), NormalizeMarkdown(document.ToMarkdown()));
    }

    private static string NormalizeMarkdown(string markdown) {
        return markdown.Replace("\r\n", "\n").TrimEnd('\n');
    }

    private static string LoadHtmlFixture(string name) {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            name);

        return File.ReadAllText(path);
    }

    private static string LoadCompatibilityFixture(string name) {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            "Compatibility",
            name);

        return File.ReadAllText(path);
    }
}
