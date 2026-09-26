using System.Collections.Concurrent;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_PrintLayoutWidthKeepsCapturedViewportForMediaQueries() {
        const string html = """
            <style>
              #medium, #large, #medium-element, #large-element { display:none }
              @media (min-width:768px) and (max-width:991px) { #medium { display:block } }
              @media (min-width:992px) { #large { display:block } }
            </style>
            <style media="(min-width:768px) and (max-width:991px)">#medium-element { display:block }</style>
            <style media="(min-width:992px)">#large-element { display:block }</style>
            <p id="medium">MediumViewport</p><p id="large">LargeViewport</p>
            <p id="medium-element">MediumStyleElement</p><p id="large-element">LargeStyleElement</p>
            """;
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            ViewportWidth = 816D,
            PrintLayoutWidthCssPixels = 1200D
        };

        string text = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(options)).ExtractText();

        Assert.Contains("MediumViewport", text, StringComparison.Ordinal);
        Assert.DoesNotContain("LargeViewport", text, StringComparison.Ordinal);
        Assert.Contains("MediumStyleElement", text, StringComparison.Ordinal);
        Assert.DoesNotContain("LargeStyleElement", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlPdf_PrintLayoutWidthSelectsViewportStylesheetAndPictureSource() {
        const string html = """
            <link rel="stylesheet" href="https://assets.example.test/medium.css" media="(min-width:768px) and (max-width:991px)">
            <link rel="stylesheet" href="https://assets.example.test/large.css" media="(min-width:992px)">
            <picture>
              <source media="(min-width:768px) and (max-width:991px)" srcset="https://assets.example.test/medium.svg" type="image/svg+xml">
              <source media="(min-width:992px)" srcset="https://assets.example.test/large.svg" type="image/svg+xml">
              <img src="https://assets.example.test/fallback.svg" width="8" height="8" alt="Fallback">
            </picture>
            """;
        var requested = new ConcurrentBag<string>();
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            HonorCssPageRules = false,
            ViewportWidth = 816D,
            PrintLayoutWidthCssPixels = 1200D,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                requested.Add(request.Uri.AbsoluteUri);
                bool stylesheet = request.Kind == HtmlResourceKind.Stylesheet;
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(stylesheet
                    ? "body { color: #123456 }"
                    : "<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'></svg>");
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(bytes,
                    stylesheet ? "text/css" : "image/svg+xml"));
            }
        };

        await HtmlConversionDocument.Parse(html).ToPdfBytesAsync(options);

        Assert.Contains("https://assets.example.test/medium.css", requested);
        Assert.Contains("https://assets.example.test/medium.svg", requested);
        Assert.DoesNotContain("https://assets.example.test/large.css", requested);
        Assert.DoesNotContain("https://assets.example.test/large.svg", requested);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthUsesNestedFrameViewportForItsMediaQueries() {
        const string html = """
            <iframe style="width:300px;height:100px" srcdoc="
              <style>
                #narrow, #wide { display:none }
                @media (max-width:400px) { #narrow { display:block } }
                @media (min-width:768px) { #wide { display:block } }
              </style>
              <p id='narrow'>NarrowFrame</p><p id='wide'>WideFrame</p>
            "></iframe>
            """;
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            ViewportWidth = 816D,
            PrintLayoutWidthCssPixels = 1200D
        };

        string text = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(options)).ExtractText();

        Assert.Contains("NarrowFrame", text, StringComparison.Ordinal);
        Assert.DoesNotContain("WideFrame", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthFitsWideRowsOnPhysicalA4() {
        const string html = """
            <div style="width:66.666%">
              <div style="display:flex;flex-wrap:wrap">
                <div style="flex:none;width:310px;height:30px;background:#ff0000">FirstCard</div>
                <div style="flex:none;width:310px;height:30px;background:#0000ff"><a href="https://example.com/second">SecondCard</a></div>
              </div>
            </div>
            """;
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(24D),
            HonorCssPageRules = false,
            PrintLayoutWidthCssPixels = 1200D
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
        (double width, double height) = read.Pages[0].GetPageSize();
        PdfCore.PdfLogicalPage page = PdfCore.PdfDocumentReadResult.Load(pdf).Pages[0];
        PdfCore.PdfLogicalTextBlock first = Assert.Single(page.TextBlocks, block => block.Text.Contains("FirstCard", StringComparison.Ordinal));
        PdfCore.PdfLogicalTextBlock second = Assert.Single(page.TextBlocks, block => block.Text.Contains("SecondCard", StringComparison.Ordinal));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.com/second"));

        Assert.Equal(OfficePageSizes.A4.WidthInches * 72D, width, 2);
        Assert.Equal(OfficePageSizes.A4.HeightInches * 72D, height, 2);
        Assert.Equal(first.BaselineY, second.BaselineY, 2);
        Assert.InRange(first.XStart, 17D, 20D);
        Assert.True(second.XStart > first.XStart + 100D);
        Assert.InRange(link.SourceLink.X1, second.XStart - 5D, width);
        Assert.InRange(link.SourceLink.X2, link.SourceLink.X1, width);
        Assert.Contains("FirstCard", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("SecondCard", read.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthRejectsNonPrintIntentAndWidthsBelowAuthoredContent() {
        var options = new HtmlToPdfOptions {
            PrintLayoutWidthCssPixels = 500D
        };
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Content</p>");

        Assert.Throws<ArgumentOutOfRangeException>(() => document.ToPdfBytes(options));

        HtmlRenderRequest screen = HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Pdf, options);
        Assert.Throws<ArgumentException>(() => document.RenderToPdfBytes(screen));
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthRetainsAuthoredSheetAndMargins() {
        string html = "<style>@page { size: 8in 10in; margin: .5in }"
            + "html,body { min-width: 1000px; margin: 0 }"
            + "div { height: 25px; line-height: 25px; font-size: 20px }</style>"
            + string.Concat(Enumerable.Range(1, 35).Select(index => "<div>Row " + index + "</div>"));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        byte[] regularPdf = document.ToPdfBytes(new HtmlToPdfOptions());
        byte[] fittedPdf = document.ToPdfBytes(new HtmlToPdfOptions {
            PrintLayoutWidthCssPixels = 1000D
        });
        PdfCore.PdfReadDocument regular = PdfCore.PdfReadDocument.Open(regularPdf);
        PdfCore.PdfReadDocument fitted = PdfCore.PdfReadDocument.Open(fittedPdf);

        Assert.Equal(2, regular.Pages.Count);
        Assert.Single(fitted.Pages);
        (double regularWidth, double regularHeight) = regular.Pages[0].GetPageSize();
        (double fittedWidth, double fittedHeight) = fitted.Pages[0].GetPageSize();
        Assert.Equal(576D, regularWidth, 1);
        Assert.Equal(720D, regularHeight, 1);
        Assert.Equal(regularWidth, fittedWidth, 1);
        Assert.Equal(regularHeight, fittedHeight, 1);
        Assert.Contains("Row 35", fitted.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthRetainsNamedPageSize() {
        const string html = """
            <style>
              @page { size: 300px 210px; margin: 10px }
              @page narrow { size: 180px 210px; margin: 10px }
              body, p, section { margin: 0 }
              section { page: narrow; break-before: page }
            </style>
            <p>Wide page</p><section>Narrow page</section>
            """;
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            PrintLayoutWidthCssPixels = 400D
        });
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);

        Assert.Equal(2, read.Pages.Count);
        Assert.Equal(225D, read.Pages[0].GetPageSize().Width, 1);
        Assert.Equal(135D, read.Pages[1].GetPageSize().Width, 1);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthRejectsOversizedNamedPage() {
        const string html = """
            <style>
              @page { size: 800px 1000px; margin: 0 }
              @page big { size: 2000px 2000px; margin: 0 }
              section { page: big; break-before: page }
            </style>
            <p>First page</p><section>Named page</section>
            """;

        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlConversionDocument.Parse(html).ToPdfBytes(
            new HtmlToPdfOptions { PrintLayoutWidthCssPixels = 20000D }));
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthKeepsPrinterMarksAtPhysicalSize() {
        const string html = """
            <style>
              @page { size: 100px 80px; margin: 10px; bleed: 4px; marks: crop cross }
              body { margin: 0 }
            </style>
            <p>Print production</p>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlPdfRenderResult regular = HtmlPdfRenderedConverter.Convert(document, new HtmlToPdfOptions());
        HtmlPdfRenderResult fitted = HtmlPdfRenderedConverter.Convert(document, new HtmlToPdfOptions {
            PrintLayoutWidthCssPixels = 160D
        });
        HtmlRenderPage regularPage = Assert.Single(regular.RenderResult!.Document.Pages);
        HtmlRenderPage fittedPage = Assert.Single(fitted.RenderResult!.Document.Pages);
        HtmlRenderShape regularMark = Assert.Single(regularPage.Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "@page marks:crop-top-left-h");
        HtmlRenderShape fittedMark = Assert.Single(fittedPage.Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "@page marks:crop-top-left-h");
        PdfCore.PdfPageInfo regularPdfPage = Assert.Single(PdfCore.PdfInspector.Inspect(regular.Document.ToBytes()).Pages);
        PdfCore.PdfPageInfo fittedPdfPage = Assert.Single(PdfCore.PdfInspector.Inspect(fitted.Document.ToBytes()).Pages);

        Assert.Equal(regularMark.Width, fittedMark.Width * 0.5D, 3);
        Assert.Equal(regularMark.Shape.StrokeWidth, fittedMark.Shape.StrokeWidth * 0.5D, 3);
        Assert.Equal(regularPdfPage.Geometry.MediaBox!.Width, fittedPdfPage.Geometry.MediaBox!.Width, 3);
        Assert.Equal(regularPdfPage.TrimBox!.Width, fittedPdfPage.TrimBox!.Width, 3);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthKeepsGalleryPdfAndPreviewOnOneRenderedPage() {
        string html = "<style>@page { size: 8in 10in; margin: .5in }"
            + "html,body { min-width: 1000px; margin: 0 }"
            + "div { height: 25px; line-height: 25px }</style>"
            + string.Concat(Enumerable.Range(1, 35).Select(index => "<div>Row " + index + "</div>"));
        var options = new HtmlRenderCapabilityGalleryOptions(
            new HtmlCapabilityGalleryScenario("authored-fit", "Authored fit", "Rendering", "Print fitting")) {
            RenderOptions = new HtmlToPdfOptions { PrintLayoutWidthCssPixels = 1000D }
        };
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.AuthoredFit." + Guid.NewGuid().ToString("N"));
        try {
            HtmlConversionDocument.Parse(html).SaveRenderCapabilityGallery(directory, options);
            PdfCore.PdfReadDocument pdf = PdfCore.PdfReadDocument.Open(File.ReadAllBytes(Path.Combine(directory, "authored-fit.pdf")));
            Assert.Single(pdf.Pages);
            Assert.Equal(576D, pdf.Pages[0].GetPageSize().Width, 1);
            Assert.Contains("Row 35", pdf.ExtractText(), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthKeepsInteractiveFieldAndFragmentTargetOnPhysicalPage() {
        const string html = """
            <a href="#target">Jump</a>
            <div style="height:200px"></div>
            <p id="target">Target</p>
            <input name="query" value="Hello" style="width:200px;height:30px">
            """;
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            PrintLayoutWidthCssPixels = 1200D
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfNamedDestination target = Assert.Single(info.NamedDestinations, destination => destination.Name == "html-fragment:target");
        PdfCore.PdfFormField field = Assert.Single(info.FormFields, item => item.Name == "query");

        Assert.InRange(target.DestinationTop!.Value, 650D, OfficePageSizes.A4.HeightInches * 72D);
        Assert.True(Assert.Single(field.Widgets).X2 <= OfficePageSizes.A4.WidthInches * 72D);
    }
}
