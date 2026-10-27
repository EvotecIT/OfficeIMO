using System.Reflection;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownAllSeverityBatch15SecurityTests {
    [Fact]
    public void TypedHtmlRenderingRejectsUnsafeSchemesButKeepsRasterDataImages() {
        var document = MarkdownDocument.Create()
            .Add(new ParagraphBlock(new InlineSequence()
                .Link("script", "java\tscript:alert(1)")
                .Image("svg", "data:image/svg+xml,<svg onload=alert(1)></svg>")
                .Image("png", "data:image/png;base64,AQID")));

        string html = document.ToHtml(new MarkdownHtmlExportOptions
        {
            Style = MarkdownHtmlStyle.Plain,
            CssDelivery = MarkdownCssDelivery.None,
            BodyClass = null,

            Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment
        });

        Assert.DoesNotContain("javascript", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("image/svg", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data:image/png;base64,AQID", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PictureSrcSetFiltersEveryCandidateAgainstImageOriginPolicy() {
        var image = new ImageBlock("https://images.example/photo.png", "photo");
        image.PictureSources.Add(new ImagePictureSource(
            "https://images.example/photo.webp",
            srcSet: "https://images.example/photo.webp 1x, https://tracker.example/pixel.webp 2x, javascript:alert(1) 3x"));

        var options = new MarkdownHtmlExportOptions();
        options.AllowedHttpImageHosts.Add("images.example");
        string html = MarkdownDocument.Create().Add(image).ToHtml(options);

        Assert.Contains("https://images.example/photo.webp 1x", html, StringComparison.Ordinal);
        Assert.DoesNotContain("tracker.example", html, StringComparison.Ordinal);
        Assert.DoesNotContain("javascript", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PictureSrcSetFiltersCommaSeparatedCandidatesWithoutWhitespace() {
        var image = new ImageBlock("https://images.example/photo.png", "photo");
        image.PictureSources.Add(new ImagePictureSource(
            "https://tracker.example/fallback.webp",
            srcSet: "https://images.example/photo.webp,https://tracker.example/pixel.webp 2x"));

        var options = new MarkdownHtmlExportOptions();
        options.AllowedHttpImageHosts.Add("images.example");
        string html = MarkdownDocument.Create().Add(image).ToHtml(options);

        Assert.Contains("https://images.example/photo.webp", html, StringComparison.Ordinal);
        Assert.DoesNotContain("tracker.example", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PictureSrcSetRejectsSlashMalformedAbsoluteHttpCandidates() {
        var image = new ImageBlock("https://images.example/photo.png", "photo");
        image.PictureSources.Add(new ImagePictureSource(
            "https://images.example/photo.webp",
            srcSet: "https://images.example/photo.webp 1x, https:/tracker.example/pixel.webp 2x"));

        var options = new MarkdownHtmlExportOptions
        {
            BlockExternalHttpImages = true,
            BaseUri = new Uri("https://images.example/")
        };
        options.AllowedHttpImageHosts.Add("images.example");
        string html = MarkdownDocument.Create().Add(image).ToHtml(options);

        Assert.Contains("https://images.example/photo.webp 1x", html, StringComparison.Ordinal);
        Assert.DoesNotContain("tracker.example", html, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkPolicyKeepsNonExecutableProtocolsButRejectsMalformedPolicyTargets() {
        var document = MarkdownDocument.Create()
            .Add(new ParagraphBlock(new InlineSequence()
                .Link("archive", "ftp://files.example/archive")
                .Link("blocked", "https:/tracker.example/pixel")));
        var options = new MarkdownHtmlExportOptions();
        options.AllowedHttpLinkHosts.Add("safe.example");

        string html = document.ToHtml(options);

        Assert.Contains("href=\"ftp://files.example/archive\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("tracker.example", html, StringComparison.Ordinal);
    }

    [Fact]
    public void InvalidThemeColorsCannotEscapeTheStyleElement() {
        string payload = "red;}</style><script>alert(1)</script>";
        string html = MarkdownDocument.Create().H1("safe").ToHtml(new MarkdownHtmlExportOptions
        {
            ColorOverrides = new MarkdownHtmlColorOverrides { AccentLight = payload },

            Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document
        });

        Assert.DoesNotContain(payload, html, StringComparison.Ordinal);
        Assert.DoesNotContain("<script>alert(1)</script>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CustomAcronymsAreHandledWithoutDynamicRegularExpressions() {
        string transformed = MarkdownHeaderTransforms.Pretty("DnsReport", new[] {
            string.Empty,
            "(",
            new string('A', 64),
            "DNS"
        });

        Assert.Equal("DNS Report", transformed);
    }

    [Fact]
    public void CallerControlledCssScopesDoNotGrowStaticCache() {
        for (int index = 0; index < 128; index++) {
            MarkdownDocument.Create().H1("safe").ToHtml(new MarkdownHtmlExportOptions
            {
                CssScopeSelector = ".tenant-" + index,

                Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document
            });
            MarkdownDocument.Create().H1("safe").ToHtml(new MarkdownHtmlExportOptions
            {
                Style = (MarkdownHtmlStyle)(10_000 + index),
                Theme = new MarkdownVisualTheme { HtmlStyle = (MarkdownHtmlStyle)(20_000 + index) },

                Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document
            });
        }

        FieldInfo field = typeof(HtmlRenderer).GetField("_unscopedBaseCssCache", BindingFlags.NonPublic | BindingFlags.Static)!;
        object cache = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<MarkdownHtmlStyle, string>>(field.GetValue(null));
        int count = (int)cache.GetType().GetProperty("Count")!.GetValue(cache)!;
        Assert.InRange(count, 1, Enum.GetValues(typeof(MarkdownHtmlStyle)).Length);
    }

    [Fact]
    public void NestedStrongNormalizationHasABoundedNumberOfPasses() {
        string adversarial = "**outer " + string.Concat(Enumerable.Repeat("**inner value** ", 2_000)) + "tail**";

        string normalized = MarkdownInputNormalizer.Normalize(adversarial, new MarkdownInputNormalizationOptions {
            NormalizeNestedStrongDelimiters = true
        });

        Assert.NotEmpty(normalized);
        Assert.InRange(normalized.Length, 1, adversarial.Length * 2);
    }
}
