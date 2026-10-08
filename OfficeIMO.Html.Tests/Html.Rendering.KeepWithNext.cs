using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPagedMedia_BreakAfterAvoidKeepsHeadingWithFollowingLine(bool keepWithNext) {
        string html = "<body style='margin:0;font:10px/10px Arial'><div style='height:40px'>Prelude</div>"
            + $"<h3 style='margin:0;height:15px;{(keepWithNext ? "break-after:avoid" : string.Empty)}'>Heading</h3>"
            + "<p style='margin:0'>First line of following paragraph.</p></body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.True(rendered.Pages.Count >= 2);
        Assert.Equal(!keepWithNext, EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>().Any(text => text.Text.Contains("Heading", StringComparison.Ordinal)));
        Assert.Equal(keepWithNext, EnumerateRenderVisuals(rendered.Pages[1].Scene)
            .OfType<HtmlRenderText>().Any(text => text.Text.Contains("Heading", StringComparison.Ordinal)));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("First", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_NestedBreakAfterAvoidKeepsHeadingWithFollowingLine() {
        const string html = "<body style='margin:0;font:10px/10px Arial'><main>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:15px;break-after:avoid'>Heading</h3>"
            + "<p style='margin:0'>First line of following paragraph.</p>"
            + "</main></body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.True(rendered.Pages.Count >= 2);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_ForcedBreakBeforeOverridesHeadingKeep() {
        const string html = "<body style='margin:0;font:10px/10px Arial'>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:15px;break-after:avoid'>Heading</h3>"
            + "<p style='margin:0;break-before:page'>First line of following paragraph.</p>"
            + "</body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("First", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_KeepIncludesUnbreakableFollowingBlock() {
        const string html = "<body style='margin:0;font:10px/10px Arial'>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:10px;break-after:avoid'>Heading</h3>"
            + "<section style='height:30px;break-inside:avoid'>Following</section>"
            + "</body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Following", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_KeepDoesNotCrossNestedForcedBreak() {
        const string html = "<body style='margin:0;font:10px/10px Arial'>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:15px;break-after:avoid'>Heading</h3>"
            + "<section><div style='height:5px'></div><p style='margin:0;break-before:page'>Following</p></section>"
            + "</body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_KeepRespectsShorterNextPage() {
        const string html = "<style>@page{size:100px 60px;margin:0}@page:left{size:100px 20px;margin:0}</style>"
            + "<body style='margin:0;font:10px/10px Arial'>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:15px;break-after:avoid'>Heading</h3>"
            + "<p style='margin:0'>Following</p>"
            + "</body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = true,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_KeepUsesNextPageWithDifferentWidth() {
        const string html = "<style>@page{size:100px 60px;margin:0}@page:left{size:80px 60px;margin:0}</style>"
            + "<body style='margin:0;font:10px/10px Arial'>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:15px;break-after:avoid'>Heading</h3>"
            + "<p style='margin:0'>Following</p>"
            + "</body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = true,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Following", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPagedMedia_KeepMeasuresViewportSizedBlockOnNextPage() {
        const string html = "<style>@page{size:80px 60px;margin:0}@page:left{size:120px 30px;margin:0}</style>"
            + "<body style='margin:0;font:10px/10px Arial'>"
            + "<div style='height:40px'>Prelude</div>"
            + "<h3 style='margin:0;height:20vw;break-after:avoid'>Heading</h3>"
            + "<p style='margin:0'>Following</p>"
            + "</body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(80D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = true,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Heading", StringComparison.Ordinal));
    }
}
