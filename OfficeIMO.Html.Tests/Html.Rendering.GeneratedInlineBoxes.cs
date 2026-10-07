using OfficeIMO.Drawing;
using OfficeIMO.Html;
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
}
