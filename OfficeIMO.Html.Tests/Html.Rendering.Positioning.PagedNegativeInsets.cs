using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPagedRootPositioning_OmitsOnlyBoxesWhollyBeforeThePageArea() {
        const string html = "<style>html,body{margin:0;font:16px Arial,sans-serif}</style>"
            + "<div id='off-top' style='position:absolute;top:-50px;left:0'>OffTopMarker</div>"
            + "<div id='off-left' style='position:absolute;top:30px;left:-180px'>OffLeftMarker</div>"
            + "<div id='partial-top' style='position:absolute;top:-10px;left:0'>PartialTopMarker</div>"
            + "<div id='transformed-in' style='position:absolute;top:-50px;left:0;transform:translateY(100px)'>TransformedVisibleMarker</div>"
            + "<div id='right-margin' style='position:absolute;top:30px;left:310px'>RightMarginMarker</div>"
            + "<div id='bottom-margin' style='position:absolute;top:320px;left:0'>BottomMarginMarker</div>"
            + "<main style='padding-top:80px'>MainFlowMarker</main>";
        var options = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(400D / HtmlRenderOptions.CssPixelsPerInch, 400D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(60D)
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.DoesNotContain("OffTopMarker", text, StringComparison.Ordinal);
        Assert.DoesNotContain("OffLeftMarker", text, StringComparison.Ordinal);
        Assert.Contains("PartialTopMarker", text, StringComparison.Ordinal);
        Assert.Contains("TransformedVisibleMarker", text, StringComparison.Ordinal);
        Assert.Contains("RightMarginMarker", text, StringComparison.Ordinal);
        Assert.Contains("BottomMarginMarker", text, StringComparison.Ordinal);
        Assert.Contains("MainFlowMarker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPagedRootPositioning_PreservesDescendantPaintedInsidePageArea() {
        const string html = "<style>html,body{margin:0;font:16px Arial,sans-serif}</style>"
            + "<div style='position:absolute;top:-50px;left:0;height:10px'>OffParentMarker"
            + "<div style='position:relative;top:70px'>VisibleChildMarker</div></div>"
            + "<main style='padding-top:80px'>MainFlowMarker</main>";
        var options = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(400D / HtmlRenderOptions.CssPixelsPerInch, 400D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(60D)
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("VisibleChildMarker", text, StringComparison.Ordinal);
        Assert.Contains("MainFlowMarker", text, StringComparison.Ordinal);
    }
}
