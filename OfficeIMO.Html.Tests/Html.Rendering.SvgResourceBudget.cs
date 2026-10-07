using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlImages_SvgForeignObjectDoesNotEraseEarlierCssImageLoss(bool asyncRender) {
        const string css = ".used{background-image:url(used.png);width:20px;height:20px}";
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='40'>"
            + "<foreignObject width='80' height='40'><div xmlns='http://www.w3.org/1999/xhtml'>NESTED</div>"
            + "</foreignObject></svg>";
        var source = HtmlConversionDocument.Parse("<link rel='stylesheet' href='https://example.test/site.css'>"
            + "<div class='used'>Visible</div><img src='data:image/svg+xml;base64,"
            + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>");
        var options = new HtmlRenderOptions {
            AllowSystemFontFallback = false,
            ResourceResolver = (request, _) => Task.FromResult<HtmlResolvedResource?>(
                request.Uri.AbsolutePath.EndsWith("site.css", StringComparison.Ordinal)
                    ? new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css") : null),
            SynchronousResourceResolver = (HtmlRenderResourceRequest request, CancellationToken _, out HtmlResolvedResource? resource) => {
                resource = request.Uri.AbsolutePath.EndsWith("site.css", StringComparison.Ordinal)
                    ? new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css") : null;
                return true;
            }
        };
        HtmlRenderDocument result = asyncRender ? await HtmlRenderEngine.RenderAsync(source, options) : HtmlRenderEngine.Render(source, options);
        HtmlDiagnostic missing = Assert.Single(result.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable && item.Source == "used.png");
        Assert.Equal(HtmlDiagnosticSeverity.Warning, missing.Severity);
        Assert.Equal(OfficeConversionLossKind.Omission, missing.LossKind);
        Assert.Throws<HtmlConversionException>(() => result.RequireNoLoss());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlImages_SvgForeignObjectsShareDecodedFontBudget(bool asyncRender) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        string rule = "<style>@font-face{font-family:Probe;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(font) + "')}p{font-family:Probe}</style><p>A</p>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='40'>"
            + "<foreignObject width='80' height='40'><div xmlns='http://www.w3.org/1999/xhtml'>"
            + rule + "</div></foreignObject></svg>";
        string image = "<img src='data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>";
        var options = new HtmlRenderOptions {
            AllowSystemFontFallback = false,
            MaxResourceBytes = Math.Max(font.Length, Encoding.UTF8.GetByteCount(svg)),
            MaxTotalResourceBytes = font.Length * 3L + Encoding.UTF8.GetByteCount(svg) + 1L
        };
        var source = HtmlConversionDocument.Parse(rule + image + image);
        HtmlRenderDocument result = asyncRender ? await HtmlRenderEngine.RenderAsync(source, options) : HtmlRenderEngine.Render(source, options);
        Assert.Single(result.Fonts.Faces);
        Assert.Contains(result.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public async Task HtmlImages_SvgForeignObjectsShareResourceCountWithoutCallingResolvers(bool asyncRender, int limit) {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='40'>"
            + "<foreignObject width='100' height='40'><div xmlns='http://www.w3.org/1999/xhtml'>"
            + "<img src='data:image/png;base64," + Convert.ToBase64String(png) + "'/>"
            + "<img src='https://example.test/not-prefetched.png'/>NESTED</div></foreignObject></svg>";
        var source = HtmlConversionDocument.Parse("<img src='data:image/svg+xml;base64,"
            + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>");
        int calls = 0;
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            MaxResourceCount = limit,
            ResourceResolver = (_, _) => { calls++; return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(png, "image/png")); },
            SynchronousResourceResolver = (HtmlRenderResourceRequest _, CancellationToken _, out HtmlResolvedResource? resource) => {
                calls++; resource = new HtmlResolvedResource(png, "image/png"); return true;
            }
        };
        HtmlRenderRequest request = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options);
        HtmlPdfRenderRequestResult result = asyncRender ? await source.RenderToPdfResultAsync(request) : source.RenderToPdfResult(request);
        Assert.Equal(0, calls);
        Assert.Equal(limit == 1, result.RenderResult.Diagnostics.Any(item => item.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded));
        byte[] pdf = result.ToBytes();
        Assert.Equal(limit == 2 ? 1 : 0, OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(pdf).Count);
        Assert.Contains("NESTED", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }
}
