using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_DataImagesHonorTheConfiguredUrlPolicy() {
        string png = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<img id='blocked' src='data:image/png;base64," + png + "'>"
            + "<div id='blocked-background' style=\"width:10px;height:10px;background-image:url('data:image/png;base64," + png + "')\"></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Visuals).ToList();

        Assert.Empty(visuals.OfType<HtmlRenderImage>());
        Assert.Empty(visuals.OfType<HtmlRenderImagePattern>());
        Assert.Empty(visuals.OfType<HtmlRenderDrawing>());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy");
    }

    [Fact]
    public void HtmlRender_ImageDataUrisHonorPerResourceBudgetsBeforeDecoding() {
        string png = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(8, 8));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<img id='oversized' src='data:image/png;base64," + png + "'>",
            new HtmlRenderOptions {
                MaxResourceBytes = 16L,
                MaxTotalResourceBytes = 1024L,
                ViewportWidth = 40D,
                Margins = HtmlRenderMargins.All(0D)
            });

        Assert.Empty(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded);
    }

    [Fact]
    public void HtmlRender_FontAndImageDataUrisShareOneOperationBudget() {
        byte[] font = CreateHtmlRenderTestFont();
        string png = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<style>@font-face{font-family:BudgetFont;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(font)
            + "')}p{font-family:BudgetFont}</style><p>Font marker</p><img id='over-count' src='data:image/png;base64,"
            + png
            + "'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            MaxResourceCount = 1,
            MaxResourceBytes = 1024L * 1024L,
            MaxTotalResourceBytes = 2L * 1024L * 1024L,
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(rendered.Fonts.Faces);
        Assert.Empty(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded);
    }

    [Fact]
    public async Task HtmlRender_ResolverSvgMediaTypesWithParametersStayVectorForImagesAndBackgrounds() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><rect width='10' height='10' fill='red'/></svg>");
        var options = new HtmlRenderOptions {
            BaseUri = new Uri("https://assets.example.test/"),
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                new HtmlResolvedResource(svg, "image/svg+xml; charset=utf-8")),
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<img id='vector' src='image.svg' style='width:10px;height:10px'>"
                + "<div id='vector-background' style=\"width:10px;height:10px;background:url('background.svg') no-repeat\"></div>",
            options);
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Visuals).ToList();

        Assert.True(visuals.OfType<HtmlRenderDrawing>().Count() >= 2);
        Assert.Empty(visuals.OfType<HtmlRenderImage>());
    }
}
