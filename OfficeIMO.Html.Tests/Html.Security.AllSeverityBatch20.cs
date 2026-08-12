using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityBatch20SecurityTests {
    private const string HiddenPixel =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";
    private const string VisiblePixel =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";

    [Fact]
    public void HiddenTableImagesDoNotParticipateInIntrinsicSizingOrDecode() {
        string html = """
            <table style="width:100px;table-layout:auto">
              <tr>
                <td id="hidden"><img style="display:none" src="{{HiddenPixel}}"></td>
                <td id="visible"><img src="{{VisiblePixel}}"></td>
              </tr>
            </table>
            """
            .Replace("{{HiddenPixel}}", HiddenPixel)
            .Replace("{{VisiblePixel}}", VisiblePixel);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            html,
            new HtmlRenderOptions {
                ViewportWidth = 120D,
                ViewportHeight = 40D,
                Margins = HtmlRenderMargins.All(0D),
                MaxResourceCount = 1
            });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(DecodeDataUri(VisiblePixel), image.Bytes);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded);
    }

    [Fact]
    public void ImagesInsideHiddenTableDescendantsDoNotParticipateInIntrinsicSizingOrDecode() {
        string html = """
            <table style="width:100px;table-layout:auto">
              <tr>
                <td id="hidden"><div style="display:none"><img src="{{HiddenPixel}}"></div></td>
                <td id="visible"><img src="{{VisiblePixel}}"></td>
              </tr>
            </table>
            """
            .Replace("{{HiddenPixel}}", HiddenPixel)
            .Replace("{{VisiblePixel}}", VisiblePixel);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            html,
            new HtmlRenderOptions {
                ViewportWidth = 120D,
                ViewportHeight = 40D,
                Margins = HtmlRenderMargins.All(0D),
                MaxResourceCount = 1
            });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(DecodeDataUri(VisiblePixel), image.Bytes);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded);
    }

    [Fact]
    public void InvalidTwoValueBorderSpacingDoesNotPartiallyApplyFirstToken() {
        Assert.False(HtmlCssTableParser.TryParseBorderSpacing(
            "4px invalid",
            fontSize: 16D,
            rootFontSize: 16D,
            out double horizontal,
            out double vertical));

        Assert.Equal(0D, horizontal);
        Assert.Equal(0D, vertical);
    }

    [Fact]
    public void PaginatedTableRetainsBottomCaptionAfterTrailingFooter() {
        const string html = """
            <table style="width:100px;margin:0;caption-side:bottom;font-size:8px;line-height:12px">
              <caption>BottomCaption</caption>
              <tbody>
                <tr><td>BodyOne</td></tr>
                <tr><td>BodyTwo</td></tr>
                <tr><td>BodyThree</td></tr>
                <tr><td>BodyFour</td></tr>
                <tr><td>BodyFive</td></tr>
                <tr><td>BodySix</td></tr>
              </tbody>
              <tfoot><tr><td>TableFooter</td></tr></tfoot>
            </table>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 48D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        int captionPage = Assert.Single(
            rendered.Pages.SelectMany((page, index) =>
                page.Visuals.OfType<HtmlRenderText>()
                    .Where(text => text.Text == "BottomCaption")
                    .Select(_ => index)));
        int[] footerPages = rendered.Pages.SelectMany((page, index) =>
                page.Visuals.OfType<HtmlRenderText>()
                    .Where(text => text.Text == "TableFooter")
                    .Select(_ => index))
            .ToArray();

        Assert.True(rendered.Pages.Count > 1);
        Assert.Equal(captionPage - 1, footerPages[footerPages.Length - 1]);
        Assert.Equal(rendered.Pages.Count - 1, captionPage);
    }

    private static byte[] DecodeDataUri(string value) {
        int separator = value.IndexOf(',');
        return Convert.FromBase64String(value.Substring(separator + 1));
    }
}
