using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlRenderingTests {
    [Fact]
    public async Task HtmlRenderAsync_UsesCallerResolverForPolicyApprovedExternalImages() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(10, 6);
        int calls = 0;
        var options = new HtmlImageExportOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                calls++;
                Assert.Equal(new Uri("https://assets.example.test/logo.png"), request.Uri);
                Assert.Equal(HtmlResourceKind.Image, request.Kind);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(imageBytes, "image/png"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(
            "<p>Resolved image</p><img src='https://assets.example.test/logo.png' width='50' height='30' alt='logo'>",
            options);
        string svg = await "<img src='https://assets.example.test/logo.png' width='50' height='30' alt='logo'>".ToSvgAsync(options);

        Assert.Equal(2, calls);
        Assert.Contains(rendered.Pages[0].Visuals, visual => visual is HtmlRenderImage image && image.ContentType == "image/png" && image.Bytes.Length == imageBytes.Length);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.Contains("data:image/png;base64", svg, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlRenderAsync_ReportsResolverTimeoutAndHonorsCallerCancellation() {
        var timeoutOptions = new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceTimeout = TimeSpan.FromMilliseconds(20D),
            ResourceResolver = async (request, cancellationToken) => {
                await Task.Delay(Timeout.Infinite, cancellationToken);
                return null;
            }
        };

        HtmlRenderDocument timedOut = await HtmlRenderEngine.RenderAsync("<img src='https://assets.example.test/slow.png' alt='slow'>", timeoutOptions);

        Assert.Contains(timedOut.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceTimeout);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => HtmlRenderEngine.RenderAsync("<p>Cancelled</p>", timeoutOptions, cancellation.Token));
    }

    [Fact]
    public async Task HtmlPdf_RenderedProfileAsync_ResolvesExternalImageAndWritesSearchablePdf() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        HtmlPdfSaveOptions options = HtmlPdfSaveOptions.CreateRenderedProfile();
        options.RenderOptions!.PageSize = new OfficePageSize(4D, 3D);
        options.RenderOptions.Margins = HtmlRenderMargins.All(16D);
        options.RenderOptions.ResourceResolver = (request, cancellationToken) =>
            Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(imageBytes, "image/png"));

        byte[] pdf = await "<h1>AsyncPdfMarker</h1><img src='https://assets.example.test/async.png' width='40' height='25' alt='async image'>".SaveAsPdfAsync(options);

        Assert.Contains("AsyncPdfMarker", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(options.ConversionReport.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
    }

    [Fact]
    public void HtmlPdf_RenderedProfile_ExposesSharedRenderResourcePolicy() {
        HtmlPdfSaveOptions options = HtmlPdfSaveOptions.CreateRenderedProfile();
        options.RenderOptions!.ResourceTimeout = TimeSpan.FromSeconds(5D);
        options.RenderOptions.MaxResourceBytes = 1024L;
        options.RenderOptions.MaxTotalResourceBytes = 4096L;
        options.RenderOptions.UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        options.RenderOptions.ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(null);

        HtmlPdfResourcePolicySummary summary = options.GetResourcePolicySummary();

        Assert.Equal(HtmlPdfProfile.Rendered, summary.Profile);
        Assert.True(summary.UsesHtmlRenderPolicy);
        Assert.False(summary.UsesWordHtmlPolicy);
        Assert.True(summary.HasRenderResourceResolver);
        Assert.Equal(TimeSpan.FromSeconds(5D), summary.RenderResourceTimeout);
        Assert.Equal(1024L, summary.RenderMaxResourceBytes);
        Assert.Equal(4096L, summary.RenderMaxTotalResourceBytes);
        Assert.Contains("https", summary.RenderAllowedUrlSchemes);
    }

    [Fact]
    public void HtmlComputedStyles_ResolveInheritedCustomPropertiesFallbacksAndCyclesForRendering() {
        string html = """
            <style>
              :root { --brand:#123456; --pad:7px; --cycle-a:var(--cycle-b); --cycle-b:var(--cycle-a); }
              .card { color:var(--brand); padding:var(--pad); background-color:var(--missing,#eeeeee); }
              .fallback { color:var(--cycle-a,#010203); }
            </style>
            <div class="card">Brand marker</div>
            <p class="fallback">Fallback marker</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(10D)
        });
        var parsed = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(parsed);

        HtmlRenderText brand = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Brand", StringComparison.Ordinal));
        HtmlRenderText fallback = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Fallback", StringComparison.Ordinal));
        HtmlRenderShape card = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source != null && shape.Source.Contains("div.card", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56), brand.Color);
        Assert.Equal(OfficeColor.FromRgb(0x01, 0x02, 0x03), fallback.Color);
        Assert.Equal(OfficeColor.FromRgb(0xEE, 0xEE, 0xEE), card.Shape.FillColor);
        Assert.Equal("7px", styles[parsed.QuerySelector(".card")!].GetValue("padding"));
    }

    [Fact]
    public void HtmlRender_Continuous_ProducesTypedVisualsWithScreenMediaAndLinks() {
        const string linkUri = "https://example.test/rendered";
        string html = """
            <style>
              .card { background-color:#123456; border:2px solid #345678; padding:8px; color:white; }
              .mode { color:#008000; }
              @media print { .mode { color:#cc0000; } }
            </style>
            <article class="card">
              <h1>Direct <a href="https://example.test/rendered">rendering</a></h1>
              <p class="mode">Screen contract</p>
            </article>
            """;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 360D,
            Margins = HtmlRenderMargins.All(12D)
        });

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Equal(HtmlRenderMode.Continuous, rendered.Mode);
        Assert.True(page.Height > 0D);
        Assert.Contains(page.Visuals, visual => visual is HtmlRenderShape shape && shape.Source != null && shape.Source.Contains("article.card", StringComparison.Ordinal));
        Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Direct", StringComparison.Ordinal));
        Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.LinkUri == linkUri && text.Text.Contains("rendering", StringComparison.Ordinal));
        Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Screen", StringComparison.Ordinal) && text.Color.G > text.Color.R);
        Assert.Contains("rendering", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlRender_Paged_UsesPrintMediaAndExplicitPageBreaks() {
        string html = """
            <style>
              .mode { color:#008000; }
              @media print { .mode { color:#cc0000; } }
            </style>
            <p class="mode">First page marker</p>
            <section style="break-before:page"><p>Second page marker</p></section>
            """;
        var options = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(20D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        IReadOnlyList<OfficeImageExportResult> images = html.ExportImages(OfficeImageExportFormat.Svg, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(2, images.Count);
        Assert.All(images, image => Assert.Equal(384, image.Width));
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("First", StringComparison.Ordinal) && text.Color.R > text.Color.G);
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Second", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRender_Paged_HonorsGenericCssPageSizeOrientationAndMargins() {
        string html = "<style>@page { size:5in 3in; margin:0.25in; }</style><p>Page rule marker</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText text = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), visual => visual.Text.Contains("Page", StringComparison.Ordinal));

        Assert.Equal(480D, page.Width, 3);
        Assert.Equal(288D, page.Height, 3);
        Assert.InRange(text.X, 23.9D, 24.1D);

        options.HonorCssPageRules = false;
        HtmlRenderDocument ignored = HtmlRenderEngine.Render(html, options);
        Assert.Equal(384D, ignored.Pages[0].Width, 3);
        Assert.Equal(384D, ignored.Pages[0].Height, 3);
    }

    [Fact]
    public void HtmlRender_Paged_AppliesPageSizeOnlyFromPrintApplicableMedia() {
        string html = "<style>@media print { @page { size: 5in 3in; } } @media screen { @page { size: 2in 2in; } }</style><p>Print page</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(12D)
        };

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(html, options).Pages);

        Assert.Equal(480D, page.Width, 3);
        Assert.Equal(288D, page.Height, 3);
    }

    [Fact]
    public void HtmlRender_Paged_FragmentsLongTextAndTablesAtStableLineAndRowBoundaries() {
        string paragraph = string.Join(" ", Enumerable.Range(0, 90).Select(index => "word" + index.ToString("D3")));
        string rows = string.Join(string.Empty, Enumerable.Range(0, 18).Select(index => "<tr><td>Row" + index.ToString("D2") + "</td><td>Value" + index.ToString("D2") + "</td></tr>"));
        string html = "<p style='background:#eef4ff;border:1px solid #446688;padding:4px'>" + paragraph + "</p><table style='border:1px solid #333'>" + rows + "</table>";
        var renderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2.5D),
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, renderOptions);
        string renderedText = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text));

        Assert.True(rendered.Pages.Count >= 3);
        Assert.Contains("word000", renderedText, StringComparison.Ordinal);
        Assert.Contains("word089", renderedText, StringComparison.Ordinal);
        Assert.Contains("Row00", renderedText, StringComparison.Ordinal);
        Assert.Contains("Row17", renderedText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "HtmlRenderBlockExceedsPage");
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);

        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = renderOptions;
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        Assert.Contains("word089", pdfText, StringComparison.Ordinal);
        Assert.Contains("Row17", pdfText, StringComparison.Ordinal);
        Assert.Equal(rendered.Pages.Count, PdfCore.PdfInspector.Inspect(pdf).PageCount);
    }

    [Fact]
    public void HtmlImageExport_RendersPngSvgTableAndDataImageWithoutNewRuntimeDependencies() {
        string pngData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(12, 8));
        string html = $"""
            <h2>Image output</h2>
            <table style="border:1px solid #333;background-color:#f5f5f5">
              <tr><th>Item</th><th>Value</th></tr>
              <tr><td>Alpha</td><td>42</td></tr>
            </table>
            <img src="data:image/png;base64,{pngData}" width="60" height="40" alt="sample image">
            """;
        var options = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 320D,
            Margins = HtmlRenderMargins.All(10D)
        };

        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult svg = html.ExportImage(OfficeImageExportFormat.Svg, options);

        Assert.True(png.Bytes.Length > 8);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.True(OfficeImageReader.TryIdentify(png.Bytes, ".png", out OfficeImageInfo pngInfo));
        Assert.Equal(png.Width, pngInfo.Width);
        Assert.Equal(png.Height, pngInfo.Height);
        string svgText = Encoding.UTF8.GetString(svg.Bytes);
        Assert.Contains("<svg", svgText, StringComparison.Ordinal);
        Assert.Contains("Image", svgText, StringComparison.Ordinal);
        Assert.Contains("output", svgText, StringComparison.Ordinal);
        Assert.Contains("Alpha", svgText, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", svgText, StringComparison.Ordinal);
        Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPdf_RenderedProfile_UsesSharedPagedLayoutAndPreservesTextAndLink() {
        const string linkUri = "https://example.test/direct-pdf";
        string html = """
            <style>@media print { h1 { color:#224466; } }</style>
            <h1>RenderedPdfMarker</h1>
            <p><a href="https://example.test/direct-pdf">RenderedLinkMarker</a></p>
            <div style="break-before:page"><p>SecondPageMarker</p></div>
            """;
        HtmlPdfSaveOptions options = HtmlPdfSaveOptions.CreateRenderedProfile();
        options.RenderOptions!.PageSize = new OfficePageSize(4D, 3D);
        options.RenderOptions.Margins = HtmlRenderMargins.All(20D);

        byte[] pdf = html.SaveAsPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Equal(2, info.PageCount);
        Assert.Contains("RenderedPdfMarker", text, StringComparison.Ordinal);
        Assert.Contains("RenderedLinkMarker", text, StringComparison.Ordinal);
        Assert.Contains("SecondPageMarker", text, StringComparison.Ordinal);
        Assert.Contains(linkUri, info.LinkUris);
        Assert.Equal(HtmlPdfProfile.Rendered, options.Profile);
        Assert.NotNull(options.RenderDiagnostics);
        Assert.DoesNotContain(options.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRender_ReportsPendingAdvancedLayoutInsteadOfSilentlyClaimingSupport() {
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            "<div style='display:flex'><span>One</span><span>Two</span></div>",
            new HtmlRenderOptions { ViewportWidth = 240D, Margins = HtmlRenderMargins.All(8D) });

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
        Assert.Equal(HtmlDiagnosticSeverity.Warning, diagnostic.Severity);
        Assert.Contains("normal flow", diagnostic.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlRenderDiagnostics_AreAllRegisteredInThePublicCatalog() {
        Assert.All(HtmlRenderDiagnosticCodes.All, code =>
            Assert.True(HtmlDiagnosticCatalog.TryGet(code, out _), code));
    }
}
