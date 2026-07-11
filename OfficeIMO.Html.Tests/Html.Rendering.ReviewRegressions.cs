using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_RenderedProfile_SkipsEmptySemanticContainers() {
        const string html = "<p></p><table><caption></caption><tr></tr></table><p>AfterEmptyMarkup</p>";

        byte[] pdf = html.SaveAsPdf(HtmlPdfSaveOptions.CreateRenderedProfile());

        Assert.Contains("AfterEmptyMarkup", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_RenderedProfile_ForcesPagedModeForSuppliedRenderOptions() {
        string html = string.Concat(Enumerable.Range(0, 40).Select(index => "<div style='line-height:20px'>Line" + index + "</div>"));
        HtmlPdfSaveOptions options = HtmlPdfSaveOptions.CreateRenderedProfile();
        options.RenderOptions = new HtmlRenderOptions {
            PageSize = new OfficePageSize(2D, 2D),
            Margins = HtmlRenderMargins.All(8D)
        };

        byte[] pdf = html.SaveAsPdf(options);

        Assert.Equal(HtmlRenderMode.Paged, options.RenderOptions.Mode);
        Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount > 1);
    }

    [Fact]
    public void HtmlRender_HonorsDisplayNoneOnTheRenderRoot() {
        const string html = "<style>body{display:none}</style><body><p>HiddenRootMarker</p></body>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("HiddenRootMarker", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("collapse")]
    public void HtmlRender_PreservesLayoutButSuppressesInvisiblePaint(string visibility) {
        string html = "<div style='height:30px;visibility:" + visibility + "'>HiddenMarker</div><div>VisibleMarker</div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("HiddenMarker", StringComparison.Ordinal));
        HtmlRenderText visible = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("VisibleMarker", StringComparison.Ordinal));
        Assert.True(visible.Y >= 30D);
    }

    [Fact]
    public void HtmlRender_MalformedPercentEncodedImageDataUsesPlaceholderInsteadOfThrowing() {
        const string html = "<img id='malformed' src='data:image/png,%ZZ' width='20' height='10' alt='invalid'>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "img#malformed");
    }

    [Fact]
    public async Task HtmlRenderAsync_UsesRenderDimensionsForMediaGatedResources() {
        byte[] image = PdfPngTestImages.CreateRgbPng(4, 4);
        var requested = new List<Uri>();
        const string html = "<style>@media (max-width:500px){.hero{width:20px;height:20px;background-image:url('https://assets.example.test/hero.png')}}</style><div class='hero'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(image, "image/png"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, options);

        Assert.Equal(new[] { new Uri("https://assets.example.test/hero.png") }, requested);
        Assert.Contains(rendered.Pages[0].Visuals, visual => visual.Source != null && visual.Source.Contains("div.hero", StringComparison.Ordinal));
    }

    [Fact]
    public async Task HtmlRenderAsync_UsesTheActivePictureCandidate() {
        byte[] activeImage = PdfPngTestImages.CreateRgbPng(8, 6);
        var requested = new List<Uri>();
        const string html = "<picture><source media='(max-width:500px)' srcset='https://assets.example.test/active.png'><img src='https://assets.example.test/fallback.png' width='40' height='30' alt='candidate'></picture>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(activeImage, "image/png"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, options);

        Assert.Equal(new[] { new Uri("https://assets.example.test/active.png") }, requested);
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(activeImage, image.Bytes);
    }

    [Fact]
    public void HtmlTable_PaintsRowGroupAndRowBackgroundsBehindTransparentCells() {
        const string html = "<table style='border-spacing:0'><tbody id='group' style='background:#0000ff'><tr id='row' style='background:#ff0000'><td style='background:transparent'>Cell</td></tr></tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        IReadOnlyList<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().ToList();
        Assert.Contains(shapes, shape => shape.Source == "tbody#group" && shape.Shape.FillColor == OfficeColor.Blue);
        Assert.Contains(shapes, shape => shape.Source == "tr#row" && shape.Shape.FillColor == OfficeColor.Red);
    }
}
