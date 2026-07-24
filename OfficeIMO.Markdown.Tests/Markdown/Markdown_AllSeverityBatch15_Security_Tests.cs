using System.Reflection;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownAllSeverityBatch15SecurityTests {
    [Fact]
    public void TypedHtmlRenderingRejectsUnsafeSchemesButKeepsRasterDataImages() {
        var document = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence()
                .Link("script", "java\tscript:alert(1)")
                .Image("svg", "data:image/svg+xml,<svg onload=alert(1)></svg>")
                .Image("png", "data:image/png;base64,AQID")));

        string html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
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

        var options = new HtmlOptions();
        options.AllowedHttpImageHosts.Add("images.example");
        string html = MarkdownDoc.Create().Add(image).ToHtmlFragment(options);

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

        var options = new HtmlOptions();
        options.AllowedHttpImageHosts.Add("images.example");
        string html = MarkdownDoc.Create().Add(image).ToHtmlFragment(options);

        Assert.Contains("https://images.example/photo.webp", html, StringComparison.Ordinal);
        Assert.DoesNotContain("tracker.example", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PictureSrcSetRejectsSlashMalformedAbsoluteHttpCandidates() {
        var image = new ImageBlock("https://images.example/photo.png", "photo");
        image.PictureSources.Add(new ImagePictureSource(
            "https://images.example/photo.webp",
            srcSet: "https://images.example/photo.webp 1x, https:/tracker.example/pixel.webp 2x"));

        var options = new HtmlOptions {
            BlockExternalHttpImages = true,
            BaseUri = new Uri("https://images.example/")
        };
        options.AllowedHttpImageHosts.Add("images.example");
        string html = MarkdownDoc.Create().Add(image).ToHtmlFragment(options);

        Assert.Contains("https://images.example/photo.webp 1x", html, StringComparison.Ordinal);
        Assert.DoesNotContain("tracker.example", html, StringComparison.Ordinal);
    }

    [Fact]
    public void InvalidThemeColorsCannotEscapeTheStyleElement() {
        string payload = "red;}</style><script>alert(1)</script>";
        string html = MarkdownDoc.Create().H1("safe").ToHtmlDocument(new HtmlOptions {
            ColorOverrides = new MarkdownHtmlColorOverrides { AccentLight = payload }
        });

        Assert.DoesNotContain(payload, html, StringComparison.Ordinal);
        Assert.DoesNotContain("<script>alert(1)</script>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CustomAcronymsAreHandledWithoutDynamicRegularExpressions() {
        string transformed = HeaderTransforms.Pretty("DnsReport", new[] {
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
            MarkdownDoc.Create().H1("safe").ToHtmlDocument(new HtmlOptions {
                CssScopeSelector = ".tenant-" + index
            });
            MarkdownDoc.Create().H1("safe").ToHtmlDocument(new HtmlOptions {
                Style = (HtmlStyle)(10_000 + index),
                Theme = new MarkdownVisualTheme { HtmlStyle = (HtmlStyle)(20_000 + index) }
            });
        }

        FieldInfo field = typeof(HtmlRenderer).GetField("_unscopedBaseCssCache", BindingFlags.NonPublic | BindingFlags.Static)!;
        object cache = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<HtmlStyle, string>>(field.GetValue(null));
        int count = (int)cache.GetType().GetProperty("Count")!.GetValue(cache)!;
        Assert.InRange(count, 1, Enum.GetValues(typeof(HtmlStyle)).Length);
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
