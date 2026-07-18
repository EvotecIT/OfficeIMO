using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Text;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_DirectRenderer_SkipsEmptySemanticContainers() {
        const string html = "<p></p><table><caption></caption><tr></tr></table><p>AfterEmptyMarkup</p>";

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());

        Assert.Contains("AfterEmptyMarkup", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_ForcesPagedModeForSuppliedRenderOptions() {
        string html = string.Concat(Enumerable.Range(0, 40).Select(index => "<div style='line-height:20px'>Line" + index + "</div>"));
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();
        options = new HtmlPdfSaveOptions {
            PageSize = new OfficePageSize(2D, 2D),
            Margins = HtmlRenderMargins.All(8D)
        };

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal(HtmlRenderMode.Paged, options.Mode);
        Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount > 1);
    }

    [Fact]
    public void HtmlRender_HonorsDisplayNoneOnTheRenderRoot() {
        const string html = "<style>body{display:none}</style><body><p>HiddenRootMarker</p></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(HtmlConversionDocument.Parse(html), options);

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

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(new[] { new Uri("https://assets.example.test/active.png") }, requested);
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(activeImage, image.Bytes);
    }

    [Fact]
    public void HtmlTable_PaintsRowGroupAndRowBackgroundsBehindTransparentCells() {
        const string html = "<table style='border-spacing:0'><tbody id='group' style='background:#0000ff'><tr id='row' style='background:#ff0000'><td style='background:transparent'>Cell</td></tr></tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        IReadOnlyList<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().ToList();
        Assert.Contains(shapes, shape => shape.Source == "tbody#group" && shape.Shape.FillColor == OfficeColor.Blue);
        Assert.Contains(shapes, shape => shape.Source == "tr#row" && shape.Shape.FillColor == OfficeColor.Red);
    }

    [Fact]
    public async Task HtmlRenderAsync_UsesRenderDimensionsWhenExpandingStylesheetImports() {
        var requested = new List<string>();
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D),
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri.AbsoluteUri);
                string css = request.Uri.AbsolutePath.EndsWith("site.css", StringComparison.Ordinal)
                    ? "@import 'mobile.css' (max-width:500px);"
                    : ".target{color:#ff0000}";
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'><p class='target'>ImportedMarker</p>",
            options);

        Assert.Contains("https://assets.example.test/mobile.css", requested);
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), visual => visual.Text.Contains("ImportedMarker", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.Red, text.Color);
    }

    [Fact]
    public void HtmlTable_SkipsRowsWithDisplayNone() {
        const string html = "<table><tr style='display:none'><td>HiddenRowMarker</td></tr><tr><td>VisibleRowMarker</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string text = string.Concat(rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(visual => visual.Text));
        Assert.DoesNotContain("HiddenRowMarker", text, StringComparison.Ordinal);
        Assert.Contains("VisibleRowMarker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_MapsSansSerifToHelvetica() {
        const string html = "<p style='font-family:sans-serif'>SansSerifMarker</p>";

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());
        string rawPdf = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/BaseFont /Helvetica", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Times-Roman", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_ExtractsColorFromBackgroundShorthandWithImage() {
        const string html = "<div id='target' style=\"width:40px;height:20px;background:#fee url('missing.png') no-repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#target" && shape.Shape.FillColor == OfficeColor.FromRgb(255, 238, 238));
    }

    [Fact]
    public void HtmlRender_Paged_IgnoresPageRulesInsideInactiveSupports() {
        const string html = "<style>@supports (not-a-real-prop:value){@page{margin:0}}</style><p>SupportedPageMarker</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(20D)
        });

        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), visual => visual.Text.Contains("SupportedPageMarker", StringComparison.Ordinal));
        Assert.InRange(text.X, 19.9D, 20.1D);
    }

    [Fact]
    public void HtmlRender_DisplayContentsSuppressesTheElementBox() {
        const string html = "<div id='contents' style='display:contents;background:#ff0000;padding:20px'><div id='child' style='background:#0000ff'>ContentsMarker</div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#contents");
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#child" && shape.Shape.FillColor == OfficeColor.Blue);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("ContentsMarker", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRender_ListStyleNoneSuppressesListMarkers() {
        const string html = "<ul style='list-style:none'><li>MarkerlessItem</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Source == "list-marker");
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("MarkerlessItem", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRender_PictureIgnoresInactiveInlineSource() {
        byte[] inactive = PdfPngTestImages.CreateRgbPng(2, 2);
        byte[] fallback = PdfPngTestImages.CreateRgbPng(3, 3);
        string html = "<picture><source media='(max-width:1px)' src='data:image/png;base64," + Convert.ToBase64String(inactive)
            + "'><img src='data:image/png;base64," + Convert.ToBase64String(fallback) + "' width='30' height='30'></picture>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(fallback, image.Bytes);
    }
}
