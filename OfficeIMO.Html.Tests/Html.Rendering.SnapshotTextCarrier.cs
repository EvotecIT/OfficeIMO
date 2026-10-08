using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(true, false, 389)]
    [InlineData(true, true, 389)]
    [InlineData(false, false, 389)]
    [InlineData(true, false, 350)]
    public void SnapshotClippedText_KeepsOneVisibleExtractionCarrier(bool scoped, bool tracked, int top) {
        string face = scoped ? "@font-face{font-family:Proof;src:url(data:font/otf;base64,"
            + Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSansPro-Regular.otf"))) + ")}" : "";
        string html = "<style>html,body{margin:0}" + face + "p{margin:0;font:20px/30px "
            + (scoped ? "Proof" : "Courier") + ";" + (tracked ? "letter-spacing:2px" : "") + "}</style>"
            + "<div style='height:" + top + "px'></div><p>BoundaryExtractionMarker</p>"
            + "<div style='height:80px'></div><p>TrailingExtractionMarker</p>";
        var options = new HtmlToPdfOptions { ViewportWidth = 400D, ViewportHeight = 400D,
            PageSize = new OfficePageSize(400D / 96D, 400D / 96D), Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false, AllowSystemFontFallback = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options));
        PdfReadDocument pdf = PdfReadDocument.Open(result.ToBytes());
        Assert.Equal(2, pdf.Pages.Count);
        Assert.Equal(1, pdf.ExtractText().Split(new[] { "BoundaryExtractionMarker" }, StringSplitOptions.None).Length - 1);
        PdfTextSpan carrier = Assert.Single(pdf.Pages[0].GetTextSpans(), span => span.Text == "BoundaryExtractionMarker");
        Assert.InRange(carrier.Y, 0D, 300D);
        Assert.InRange(carrier.X, 0D, 300D);
    }
}
