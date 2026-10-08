using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("before")]
    [InlineData("after")]
    public void HtmlGeneratedContent_NonemptyInlineBlockRetainsPaintInsetsAndLink(string pseudo) {
        string html = "<style>body,p{margin:0}body{font-size:16px;line-height:20px}"
            + ".badge::" + pseudo + "{content:'LIVE';display:inline-block;margin-left:8px;"
            + "padding:4px 6px;border:1px solid #ff0000;background:#ffff00}</style>"
            + "<p><a href='https://example.test/badge'><span class='badge'>Label</span></a> <em>tail</em></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        string source = "span.badge::" + pseudo;
        HtmlRenderShape box = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            v => v.Source == source && v.Shape.FillColor == OfficeColor.FromRgb(255, 255, 0));
        HtmlRenderText generated = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "LIVE");
        HtmlRenderText label = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Label");
        Assert.True(generated.TextAdvanceWidth.HasValue);
        Assert.InRange(box.Width, generated.TextAdvanceWidth.Value + 13.99D, generated.TextAdvanceWidth.Value + 14.01D);
        Assert.Equal(box.X + 7D, generated.X, 3);
        Assert.Equal(box.Y + 5D, generated.Y, 3);
        Assert.Equal("https://example.test/badge", generated.LinkUri);
        Assert.True(pseudo == "before" ? label.X >= box.X + box.Width : box.X >= label.X + label.TextAdvanceWidth);
        Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "tail");
    }

    [Fact]
    public void HtmlGeneratedContent_InlineBlockRetainsExplicitBorderBoxClipAndOpacity() {
        const string html = """
            <style>body,p{margin:0}.badge::after{content:'WIDE WORDS';display:inline-block;
              font-size:10px;line-height:12px;width:32px;height:18px;box-sizing:border-box;
              padding:2px;border:1px solid #ff0000;background:#ffff00;overflow:hidden;opacity:.5}</style>
            <p><a href="https://example.test/badge"><span class="badge">Label</span></a></p>
            """;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderEffectGroup effect = Assert.Single(visuals.OfType<HtmlRenderEffectGroup>(), v => v.Source == "span.badge::after");
        Assert.Equal(.5D, effect.Opacity, 3);
        HtmlRenderShape box = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            v => v.Source == "span.badge::after" && v.Shape.FillColor == OfficeColor.FromRgb(255, 255, 0));
        Assert.Equal(32D, box.Width, 3);
        Assert.Equal(18D, box.Height, 3);
        Assert.Contains(visuals, v => v is HtmlRenderClipGroup);
        Assert.Equal(new[] { "WIDE", "WORDS" }, visuals.OfType<HtmlRenderText>()
            .Where(t => t.Source == "span.badge::after").Select(t => t.Text).ToArray());
    }

    [Fact]
    public void HtmlGeneratedContent_InlineBlockAtRootKeepsOwnOverflowClip() {
        const string html = "<style>body{margin:0;overflow:hidden}body::before{content:'OVERFLOW';"
            + "display:inline-block;width:20px;height:12px;white-space:nowrap;overflow:hidden;"
            + "background:#ffff00}p{margin:0}</style><p>following</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        Assert.Single(visuals.OfType<HtmlRenderShape>(), v => v.Source == "body::before"
            && v.Shape.FillColor == OfficeColor.FromRgb(255, 255, 0) && v.Width == 20D);
        HtmlRenderClipGroup clip = Assert.Single(visuals.OfType<HtmlRenderClipGroup>(),
            v => v.Width == 20D && EnumerateRenderVisuals(v.Visuals).OfType<HtmlRenderText>().Any(t => t.Text == "OVERFLOW"));
        Assert.Equal(20D, clip.Width, 3);
        Assert.Equal(12D, clip.Height, 3);
        Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "following");
    }

    [Fact]
    public void HtmlGeneratedContent_InlineBlockPreLinesUseLongestLineIntrinsicWidth() {
        const string html = "<style>body,p{margin:0}.badge::after{content:'LONG\\A X';white-space:pre;"
            + "display:inline-block;font-size:10px;line-height:12px;padding:2px;background:#ffff00}</style>"
            + "<p><span class='badge'>Label</span></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderShape box = Assert.Single(visuals.OfType<HtmlRenderShape>(), v => v.Source == "span.badge::after"
            && v.Shape.FillColor == OfficeColor.FromRgb(255, 255, 0));
        HtmlRenderText first = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "LONG");
        HtmlRenderText second = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "X");
        Assert.InRange(box.Width, first.TextAdvanceWidth!.Value + 3.99D, first.TextAdvanceWidth.Value + 4.01D);
        Assert.Equal(12D, second.Y - first.Y, 3);
        Assert.Equal(28D, box.Height, 3);
    }
    [Theory]
    [InlineData("before", false, "background:blue")]
    [InlineData("after", false, "background:blue")]
    [InlineData("before", true, "opacity:.5")]
    [InlineData("after", true, "opacity:.5")]
    public void HtmlGeneratedContent_InlineBoxUnderPaintedInlineOwnerKeepsText(string pseudo, bool ancestor, string ownerStyle) {
        string html = "<style>body,p{margin:0}.painted{" + ownerStyle + "}.badge::" + pseudo
            + "{content:'LIVE';display:inline-block;padding:2px;border:1px solid red;background:yellow}</style><p>"
            + (ancestor ? "<a class='painted' href='https://example.test/badge'><span class='badge'>Label</span></a>"
                : "<span class='badge painted'>Label</span>") + "</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderText generated = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "LIVE");
        Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Label");
        if (ancestor) {
            Assert.Equal("https://example.test/badge", generated.LinkUri);
            Assert.Single(visuals.OfType<HtmlRenderEffectGroup>(), g => g.Opacity == .5D);
        }
    }

    [Theory]
    [InlineData("span", false, "inline-block")]
    [InlineData("body", true, "inline-block")]
    [InlineData("body", true, "block")]
    public void HtmlGeneratedContent_InlineBoxImageKeepsIntrinsicSizeWithoutRepeatingPseudoPaint(string origin, bool mixed, string display) {
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string content = (mixed ? "'A' " : "") + "url('data:image/png;base64," + png + "')" + (mixed ? " 'B'" : "");
        string html = "<style>body,p{margin:0}" + origin + "::before{content:" + content
            + ";display:" + display + ";padding:2px;border:1px solid red;margin-left:8px;background:yellow}</style>"
            + (origin == "span" ? "<p><span>Host</span></p>" : "<p>Host</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderImage image = Assert.Single(visuals.OfType<HtmlRenderImage>());
        Assert.Equal(1D, image.Width, 3);
        Assert.Equal(1D, image.Height, 3);
        HtmlRenderShape box = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            v => v.Shape.FillColor == OfficeColor.FromRgb(255, 255, 0));
        Assert.Equal(origin + "::before", box.Source);
        Assert.InRange(image.X - box.X, 3D, box.Width - 3D);
        Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Host");
        if (mixed) {
            Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "A");
            Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "B");
        }
    }

    [Theory]
    [InlineData("before", "inline-block")]
    [InlineData("after", "inline-block")]
    [InlineData("before", "block")]
    [InlineData("after", "block")]
    public void HtmlGeneratedContent_InlineBoxPreparesArabicJoiningAndLogicalText(string pseudo, string display) {
        string html = "<style>body,p{margin:0}.badge::" + pseudo
            + "{content:'سلام';display:" + display + ";direction:rtl}</style><p><span class='badge'>Host</span></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        Assert.Contains(visuals.OfType<HtmlRenderText>().SelectMany(t => t.Text), c => c >= '\uFE70' && c <= '\uFEFF');
        Assert.Contains(visuals.OfType<HtmlRenderLogicalTextGroup>(), g => g.Text == "سلام");
        Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Host");
    }

    [Theory]
    [InlineData("inline-block")]
    [InlineData("block")]
    public void HtmlGeneratedContent_BoxUsesScopedUnicodeRangeFaces(string display) {
        string font = Convert.ToBase64String(ManagedTextShapingTestAssets.CreateFont('A', 0x05D0));
        string html = "<style>body,p{margin:0}"
            + "@font-face{font-family:Scoped;src:url('data:font/ttf;base64," + font + "');unicode-range:U+0000-007F}"
            + "@font-face{font-family:Scoped;src:url('data:font/ttf;base64," + font + "');unicode-range:U+0590-05FF}"
            + "p::before{content:'Aא';font-family:Scoped;font-size:20px;display:" + display + "}"
            + "</style><p>Host</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false
        });
        HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>()
            .Where(t => t.Source == "p::before").ToArray();
        HtmlRenderText latin = Assert.Single(text, t => t.Text == "A");
        HtmlRenderText hebrew = Assert.Single(text, t => t.Text == "א");
        Assert.StartsWith("Scoped__officeimo_", latin.Font.FamilyName, StringComparison.Ordinal);
        Assert.StartsWith("Scoped__officeimo_", hebrew.Font.FamilyName, StringComparison.Ordinal);
        Assert.NotEqual(latin.Font.FamilyName, hebrew.Font.FamilyName);
    }

    [Theory]
    [InlineData("inline-block", "image-resolution:2dppx", .5D, 1D)]
    [InlineData("block", "image-resolution:2dppx", .5D, 1D)]
    [InlineData("inline-block", "image-orientation:none", 2D, 1D)]
    [InlineData("block", "image-orientation:none", 2D, 1D)]
    public void HtmlGeneratedContent_BoxImageKeepsInheritedImageProperties(
        string display, string property, double expectedWidth, double expectedHeight) {
        var source = new OfficeRasterImage(2, 1);
        source.SetPixel(0, 0, OfficeColor.Red);
        source.SetPixel(1, 0, OfficeColor.Blue);
        byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = 100, Subsampling = OfficeJpegSubsampling.Y444,
            Metadata = new OfficeJpegMetadata(exif: CreateHtmlExifOrientation(6))
        });
        string html = "<style>body,p{margin:0}p::before{content:url('data:image/jpeg;base64,"
            + Convert.ToBase64String(jpeg) + "');display:" + display + ";" + property
            + ";padding:2px;border:1px solid red;background:yellow}</style><p>Host</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderImage image = Assert.Single(visuals.OfType<HtmlRenderImage>());
        Assert.Equal(expectedWidth, image.Width, 3);
        Assert.Equal(expectedHeight, image.Height, 3);
        Assert.Single(visuals.OfType<HtmlRenderShape>(), v => v.Shape.FillColor == OfficeColor.FromRgb(255, 255, 0));
        Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Host");
    }
}
