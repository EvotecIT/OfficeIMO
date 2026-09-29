using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlFlexRow_KeepsFittingAutoHeightLineOnNextPage() {
        const string html = """
            <style>body,p{margin:0}</style>
            <div style="height:50px">Before</div>
            <div id="row" style="display:flex;width:160px;align-items:flex-start">
              <div style="width:80px;line-height:20px">First<br>Second<br>Third<br>Fourth</div>
              <div style="width:80px;height:20px">Side</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text is "First" or "Second" or "Third" or "Fourth" or "Side");
        Assert.All(new[] { "First", "Second", "Third", "Fourth", "Side" }, marker =>
            Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == marker));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Fact]
    public void HtmlFlexWrap_KeepsFittingRowsTogetherOnNextPage() {
        const string html = """
            <style>body{margin:0}</style>
            <div style="height:50px">Before</div>
            <div style="display:flex;flex-wrap:wrap;width:160px">
              <div style="width:80px;height:35px">First</div>
              <div style="width:80px;height:35px">Second</div>
              <div style="width:80px;height:35px">Third</div>
              <div style="width:80px;height:35px">Fourth</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text is "First" or "Second" or "Third" or "Fourth");
        Assert.All(new[] { "First", "Second", "Third", "Fourth" }, marker =>
            Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == marker));
    }

    [Fact]
    public void HtmlFlexRow_PreservesNestedFittingWrappedRowBesideOversizedSibling() {
        const string html = """
            <style>body{margin:0}</style>
            <div style="height:50px">Before</div>
            <div style="display:flex;width:200px;align-items:flex-start">
              <div style="display:flex;flex-wrap:wrap;width:100px">
                <div style="width:100px;height:40px">Inner A</div>
                <div style="width:100px;height:40px">Inner B</div>
              </div>
              <div style="width:100px;line-height:20px">Tall one<br>Tall two<br>Tall three<br>Tall four<br>Tall five<br>Tall six<br>Tall seven<br>Tall eight<br>Tall nine</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2.5D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text is "Inner A" or "Inner B");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inner A");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inner B");
        Assert.Equal(1, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Count(text => text.Text == "Inner A"));
        Assert.Equal(1, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Count(text => text.Text == "Inner B"));
    }

    [Fact]
    public void HtmlFlexRow_PreservesFittingAvoidBreakItemBesideOversizedSibling() {
        const string html = """
            <style>body{margin:0}</style>
            <div style="height:50px">Before</div>
            <div style="display:flex;width:200px;align-items:flex-start">
              <div style="width:100px;break-inside:avoid">
                <div style="height:40px">Inner A</div>
                <div style="height:40px">Inner B</div>
              </div>
              <div style="width:100px;line-height:20px">Tall one<br>Tall two<br>Tall three<br>Tall four<br>Tall five<br>Tall six<br>Tall seven<br>Tall eight<br>Tall nine</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2.5D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text is "Inner A" or "Inner B");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inner A");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inner B");
    }
}
