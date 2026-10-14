using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTextFlow_UsesSharedCjkBreaksUnlessKeepAllIsRequested() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderText[] normal = RenderParagraph("width:42px", options);
        HtmlRenderText[] keep = RenderParagraph("width:42px;word-break:keep-all", options);

        Assert.True(normal.Select(text => text.Y).Distinct().Count() > 1);
        Assert.Single(keep.Select(text => text.Y).Distinct());
        Assert.Equal("日本語文書作成", string.Concat(normal.Select(text => text.Text)));
        Assert.Equal("日本語文書作成", string.Concat(keep.Select(text => text.Text)));
    }

    [Fact]
    public void HtmlTextFlow_UsesSharedPunctuationSafeBreaksBeforeEmergencyBreaking() {
        const string html = "<p id='breaks' style='margin:0;width:38px;font-family:Arial;font-size:12px;line-height:14px'>日本語、文書。作成</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        string[] lines = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .GroupBy(text => text.Y)
            .OrderBy(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(text => text.X).Select(text => text.Text)))
            .ToArray();

        Assert.True(lines.Length > 1);
        Assert.DoesNotContain(lines, line => line.StartsWith("、", StringComparison.Ordinal) || line.StartsWith("。", StringComparison.Ordinal));
        Assert.Equal("日本語、文書。作成", string.Concat(lines));
    }

    [Fact]
    public void HtmlTextFlow_KeepAllRetainsNonCjkPreferredBreaks() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderText[] text = RenderParagraph("width:48px;word-break:keep-all", options, "alpha-beta/gamma");

        Assert.True(text.Select(run => run.Y).Distinct().Count() > 1);
        Assert.Equal("alpha-beta/gamma", string.Concat(text.Select(run => run.Text)));
    }

    [Theory]
    [InlineData("\u2003", "")]
    [InlineData("\u2003", "word-break:keep-all")]
    [InlineData("\u3000", "")]
    public void HtmlTextFlow_UnicodeBreakSpaceRetainsItsGlyphAndMinContentAdvance(string separator, string additionalStyle) {
        const string wrapper = "<style>#box{margin:0;display:inline-block;width:min-content;background:yellow}</style>";
        string actual = wrapper + "<p id='box' style='" + additionalStyle + "'>alpha" + separator + "beta</p><span>after</span>";
        string control = wrapper + "<p id='box' style='white-space:nowrap'>alpha" + separator + "<br>beta</p><span>after</span>";
        HtmlRenderDocument rendered = RenderFirstLetter(actual);
        HtmlRenderVisual[] actualVisuals = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderVisual[] controlVisuals = EnumerateTextOverflowVisuals(RenderFirstLetter(control).Pages[0].Scene).ToArray();
        HtmlRenderShape actualBox = Assert.Single(actualVisuals.OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderShape controlBox = Assert.Single(controlVisuals.OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderText[] actualText = actualVisuals.OfType<HtmlRenderText>().Where(item => item.Text != "after").ToArray();
        HtmlRenderText following = Assert.Single(actualVisuals.OfType<HtmlRenderText>(), item => item.Text == "after");
        HtmlRenderText controlFollowing = Assert.Single(controlVisuals.OfType<HtmlRenderText>(), item => item.Text == "after");

        Assert.Equal("alpha" + separator + "beta", string.Concat(actualText.Select(item => item.Text)));
        Assert.Equal(2, actualText.Select(item => item.Y).Distinct().Count());
        Assert.Equal(controlBox.Width, actualBox.Width, 6);
        Assert.Equal(controlBox.Height, actualBox.Height, 6);
        Assert.Equal(controlFollowing.X, following.X, 6);
        Assert.Equal(controlFollowing.Y, following.Y, 6);
        rendered.RequireNoLoss();
    }

    private static HtmlRenderText[] RenderParagraph(string style, HtmlRenderOptions options, string text = "日本語文書作成") {
        string html = "<p style='margin:0;font-family:Arial;font-size:12px;line-height:14px;" + style + "'>" + text + "</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        return rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
    }
}
