using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Tests.Pdf;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_LossContractRejectsEveryDiagnosedFallback() {
        HtmlRenderDocument permissive = HtmlRenderTestDriver.Render(
            "<div style='transform:rotate(not-an-angle)'>Fallback</div>");

        Assert.True(permissive.HasLoss);
        HtmlConversionException documentException = Assert.Throws<HtmlConversionException>(() => permissive.RequireNoLoss());
        Assert.Equal("HTML conversion did not satisfy the required no-loss contract.", documentException.Message);
        Assert.Contains(documentException.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported);

        var strict = new HtmlRenderOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };
        HtmlConversionException renderException = Assert.Throws<HtmlConversionException>(() =>
            HtmlRenderTestDriver.Render("<div style='transform:rotate(not-an-angle)'>Fallback</div>", strict));
        Assert.Contains(renderException.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported);
    }

    [Fact]
    public async Task HtmlRenderAsync_StrictLossContractReturnsCleanOutputAndRejectsFallbacks() {
        var options = new HtmlRenderOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument clean = await HtmlRenderTestDriver.RenderAsync("<h1>Lossless static output</h1>", options);
        Assert.False(clean.HasLoss);
        Assert.Same(clean, clean.RequireNoLoss());

        await Assert.ThrowsAsync<HtmlConversionException>(() =>
            HtmlRenderTestDriver.RenderAsync("<div style='opacity:not-a-number'>Fallback</div>", options));
        Assert.Equal(HtmlRenderFidelityPolicy.RequireNoLoss, options.Clone().FidelityPolicy);
    }

    [Fact]
    public void HtmlFitWithinBoundsHighRequestedScaleBeforeSurfaceValidation() {
        OfficeImageExportResult result = HtmlConversionDocument
            .Parse("<h1>Bounded</h1><p>High requested scale, small final surface.</p>")
            .ToImage()
            .WithScale(100D)
            .FitWithin(360, 360)
            .AsPng()
            .Export();

        Assert.True(result.Width <= 360);
        Assert.True(result.Height <= 360);
    }

    [Fact]
    public async Task HtmlAsyncBatchPopulatesKnownSequenceCountBeforeEmission() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<h1>First</h1><section style='break-before:page'><h2>Second</h2></section>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D)
        };

        IReadOnlyList<OfficeImageExportResult> results = await document.ExportImagesAsync(
            OfficeImageExportFormat.Svg,
            options);

        Assert.True(results.Count >= 2);
        Assert.Equal(
            Enumerable.Range(0, results.Count).Select(index => (int?)index),
            results.Select(result => result.SequenceIndex));
        Assert.All(results, result => Assert.Equal(results.Count, result.SequenceCount));
    }

    [Fact]
    public async Task HtmlRenderAsync_UsesCallerResolverForPolicyApprovedExternalImages() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(10, 6);
        int calls = 0;
        var options = new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<p>Resolved image</p><img src='https://assets.example.test/logo.png' width='50' height='30' alt='logo'>",
            options);
        string svg = await HtmlConversionDocument.Parse("<img src='https://assets.example.test/logo.png' width='50' height='30' alt='logo'>").ToSvgAsync(options);

        Assert.Equal(2, calls);
        Assert.Contains(rendered.Pages[0].Visuals, visual => visual is HtmlRenderImage image && image.ContentType == "image/png" && image.Bytes.Length == imageBytes.Length);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.Contains("data:image/png;base64", svg, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlRenderAsync_AppliesExternalStylesheetInCascadeOrder() {
        const string stylesheet = "@page { size:4in 3in; margin:12px; } .external { color:#123456; font-family:\"Definitely Missing\", Arial, sans-serif; }";
        int calls = 0;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                calls++;
                Assert.Equal(new Uri("https://assets.example.test/theme.css"), request.Uri);
                Assert.Equal(HtmlResourceKind.Stylesheet, request.Kind);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes(stylesheet), "text/css; charset=utf-8"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/theme.css'><style>.override { color:#654321; }</style><p class='external'>External sheet</p><p class='external override'>Cascade override</p>",
            options);

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText external = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("External sheet", StringComparison.Ordinal));
        HtmlRenderText overridden = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Cascade override", StringComparison.Ordinal));
        Assert.Equal(1, calls);
        Assert.Equal(384D, page.Width, 3);
        Assert.Equal(288D, page.Height, 3);
        Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56), external.Color);
        Assert.Equal(OfficeColor.FromRgb(0x65, 0x43, 0x21), overridden.Color);
        Assert.Contains("Definitely Missing", external.Font.FamilyName, StringComparison.Ordinal);
        Assert.Contains("Arial", external.Font.FamilyName, StringComparison.Ordinal);
        Assert.Contains(",", external.Font.FamilyName, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending);
    }

    [Fact]
    public async Task HtmlRenderAsync_ResolvesRecursiveStylesheetImports() {
        var stylesheets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["https://assets.example.test/css/top.css"] = "@import 'base.css'; .top { color:#112233; }",
            ["https://assets.example.test/css/base.css"] = "@import url('palette.css') screen; .base { color:#334455; font-family:\"Missing\", Arial, sans-serif; }",
            ["https://assets.example.test/css/palette.css"] = ".palette { color:#556677; }"
        };
        var requested = new List<string>();
        var options = new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) => {
                requested.Add(request.Uri.AbsoluteUri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    System.Text.Encoding.UTF8.GetBytes(stylesheets[request.Uri.AbsoluteUri]),
                    "text/css"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/css/top.css'><p class='top'>Top import</p><p class='base'>Base import</p><p class='palette'>Palette import</p>",
            options);

        HtmlRenderText top = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Top import", StringComparison.Ordinal));
        HtmlRenderText baseText = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Base import", StringComparison.Ordinal));
        HtmlRenderText palette = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Palette import", StringComparison.Ordinal));
        Assert.Equal(new[] {
            "https://assets.example.test/css/top.css",
            "https://assets.example.test/css/base.css",
            "https://assets.example.test/css/palette.css"
        }, requested);
        Assert.Equal(OfficeColor.FromRgb(0x11, 0x22, 0x33), top.Color);
        Assert.Equal(OfficeColor.FromRgb(0x33, 0x44, 0x55), baseText.Color);
        Assert.Equal(OfficeColor.FromRgb(0x55, 0x66, 0x77), palette.Color);
        Assert.Contains("Arial", baseText.Font.FamilyName, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetImportCycle);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending);
    }

    [Fact]
    public async Task HtmlRenderAsync_EnforcesSharedCssByteLimitsAcrossResolvedStylesheets() {
        const string firstCss = ".first { color:red; }";
        const string secondCss = ".second { color:blue; }";
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = 24;
        limits.MaxTotalCssBytes = 30;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/first.css'>" +
            "<link rel='stylesheet' href='https://assets.example.test/second.css'>" +
            "<p class='first'>First</p><p class='second'>Second</p>",
            new HtmlConversionDocumentOptions { Limits = limits });
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                new HtmlResolvedResource(
                    System.Text.Encoding.UTF8.GetBytes(request.Uri.AbsolutePath.Contains("first", StringComparison.Ordinal) ? firstCss : secondCss),
                    "text/css"))
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded
            && diagnostic.Source == "https://assets.example.test/second.css");
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("First", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(255, 0, 0));
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Second", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(0, 0, 255));
    }

    [Fact]
    public async Task HtmlRenderAsync_ChargesEncodedBytesForResolvedStylesheets() {
        const string css = ".target { color:red; }";
        byte[] encodedCss = System.Text.Encoding.Unicode.GetPreamble()
            .Concat(System.Text.Encoding.Unicode.GetBytes(css))
            .ToArray();
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = System.Text.Encoding.UTF8.GetByteCount(css);
        limits.MaxTotalCssBytes = 128;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/utf16.css'><p class='target'>Text</p>",
            new HtmlConversionDocumentOptions { Limits = limits });
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                new HtmlResolvedResource(encodedCss, "text/css; charset=utf-16"))
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.CssSizeLimitExceeded
            && diagnostic.Source == "https://assets.example.test/utf16.css");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text.Contains("Text", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(255, 0, 0));
    }

    [Fact]
    public async Task HtmlRenderAsync_RejectsOversizedStylesheetBeforeQueuingImports() {
        const string oversizedCss = "@import 'must-not-load.css'; .oversized { color: red; padding: 20px; }";
        var requested = new List<string>();
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = 32;
        limits.MaxTotalCssBytes = 128;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/top.css'><p class='oversized'>Text</p>",
            new HtmlConversionDocumentOptions { Limits = limits });
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => {
                requested.Add(request.Uri.AbsoluteUri);
                string css = request.Uri.AbsolutePath.EndsWith("top.css", StringComparison.Ordinal)
                    ? oversizedCss
                    : ".unexpected { color: blue; }";
                return Task.FromResult<HtmlResolvedResource?>(
                    new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes(css), "text/css"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Equal(new[] { "https://assets.example.test/top.css" }, requested);
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.CssSizeLimitExceeded
            && diagnostic.Source == "https://assets.example.test/top.css");
    }

    [Fact]
    public async Task HtmlRenderAsync_EnforcesPerSheetCssByteLimitOnResolvedImports() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = 24;
        limits.MaxTotalCssBytes = 64;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/top.css'><p class='base'>Base</p>",
            new HtmlConversionDocumentOptions { Limits = limits });
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                new HtmlResolvedResource(
                    System.Text.Encoding.UTF8.GetBytes(
                        request.Uri.AbsolutePath.Contains("top", StringComparison.Ordinal)
                            ? "@import 'base.css';"
                            : ".base { color:rgb(1,2,3); padding:1px; }"),
                    "text/css"))
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.CssSizeLimitExceeded
            && diagnostic.Source == "base.css");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Base", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(1, 2, 3));
    }

    [Fact]
    public async Task HtmlRenderAsync_SuppressesStylesheetImportCycles() {
        var stylesheets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["https://assets.example.test/css/top.css"] = "@import 'base.css'; .cycle-top { color:#123456; }",
            ["https://assets.example.test/css/base.css"] = "@import 'top.css'; .cycle-base { color:#654321; }"
        };
        int calls = 0;
        var options = new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) => {
                calls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    System.Text.Encoding.UTF8.GetBytes(stylesheets[request.Uri.AbsoluteUri]),
                    "text/css"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/css/top.css'><p class='cycle-top'>Cycle top</p><p class='cycle-base'>Cycle base</p>",
            options);

        Assert.Equal(2, calls);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Cycle top", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(0x12, 0x34, 0x56));
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Cycle base", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(0x65, 0x43, 0x21));
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetImportCycle);
    }

    [Fact]
    public async Task HtmlRenderAsync_EnforcesStylesheetImportDepthAndResourceCountLimits() {
        var stylesheets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["https://assets.example.test/css/top.css"] = "@import 'base.css'; .top-limit { color:#112233; }",
            ["https://assets.example.test/css/base.css"] = "@import 'deep.css'; .base-limit { color:#334455; }",
            ["https://assets.example.test/css/deep.css"] = ".deep-limit { color:#556677; }"
        };
        HtmlRenderResourceResolver resolver = (request, cancellationToken) =>
            Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                System.Text.Encoding.UTF8.GetBytes(stylesheets[request.Uri.AbsoluteUri]),
                "text/css"));
        const string html = "<link rel='stylesheet' href='https://assets.example.test/css/top.css'><p class='base-limit'>Limited import</p>";

        HtmlRenderDocument depthLimited = await HtmlRenderTestDriver.RenderAsync(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(8D),
            MaxStylesheetImportDepth = 1,
            ResourceResolver = resolver
        });
        HtmlRenderDocument countLimited = await HtmlRenderTestDriver.RenderAsync(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(8D),
            MaxResourceCount = 1,
            ResourceResolver = resolver
        });

        Assert.Contains(depthLimited.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetImportDepthExceeded);
        Assert.Contains(depthLimited.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Limited import", StringComparison.Ordinal) && text.Color == OfficeColor.FromRgb(0x33, 0x44, 0x55));
        Assert.Contains(countLimited.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded);
    }

    [Fact]
    public void HtmlRender_ReportsExternalStylesheetPendingForSynchronousRendering() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<link rel='stylesheet' href='https://assets.example.test/theme.css'><p>Pending sheet</p>",
            new HtmlRenderOptions { ViewportWidth = 240D, Margins = HtmlRenderMargins.All(8D) });

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending
            && diagnostic.Source == "https://assets.example.test/theme.css");
    }

    [Fact]
    public async Task HtmlRenderAsync_RejectsNonCssStylesheetContent() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) =>
                Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes(".unsafe { color:red; }"), "text/html"))
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/not-css'><p class='unsafe'>Rejected sheet</p>",
            options);

        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), visual => visual.Text.Contains("Rejected sheet", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.Black, text.Color);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceContentTypeRejected);
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

        HtmlRenderDocument timedOut = await HtmlRenderTestDriver.RenderAsync("<img src='https://assets.example.test/slow.png' alt='slow'>", timeoutOptions);

        Assert.Contains(timedOut.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceTimeout);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => HtmlRenderTestDriver.RenderAsync("<p>Cancelled</p>", timeoutOptions, cancellation.Token));
    }

    [Fact]
    public async Task HtmlRenderAsync_CancelsLargeRenderOperation() {
        string html = "<main>" + string.Concat(Enumerable.Repeat("<div><span>Cancellation marker</span></div>", 20000)) + "</main>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => {
            Thread.Sleep(1);
            cancellation.Cancel();
        }) {
            IsBackground = true
        };
        cancellationThread.Start();

        try {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                HtmlRenderTestDriver.RenderAsync(
                    document,
                    new HtmlRenderOptions {
                        ViewportWidth = 240D,
                        MaxSurfaceHeight = 1_000_000
                    },
                    cancellation.Token));
        } finally {
            Assert.True(cancellationThread.Join(TimeSpan.FromSeconds(5D)));
        }
    }

    [Fact]
    public void HtmlRenderPage_CreateDrawingHonorsCancellation() {
        HtmlRenderPage page = HtmlRenderTestDriver.Render("<p>Drawing cancellation marker</p>").Pages[0];
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => page.CreateDrawing(cancellation.Token));
    }

    [Fact]
    public async Task HtmlImageAndRenderedPdfAsync_HonorCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            HtmlConversionDocument.Parse("<p>Image cancellation marker</p>").ExportImagesAsync(OfficeImageExportFormat.Png, cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            OfficeIMO.Html.HtmlConversionDocument.Parse("<p>PDF cancellation marker</p>").ToPdfAsync(new HtmlPdfSaveOptions(), cancellation.Token));
    }

    [Fact]
    public async Task HtmlPdf_DirectRendererAsync_ResolvesExternalImageAndWritesSearchablePdf() {
        const string html = "<h1>AsyncPdfMarker</h1><img src='https://assets.example.test/async.png' width='40' height='25' alt='async image'>";
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
        };
        options.PageSize = new OfficePageSize(4D, 3D);
        options.Margins = HtmlRenderMargins.All(16D);
        options.ResourceResolver = (request, cancellationToken) =>
            Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(imageBytes, "image/png"));

        PdfCore.PdfDocumentConversionResult result = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResultAsync(options);
        byte[] pdf = result.ToBytes();

        Assert.Contains("AsyncPdfMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
    }

    [Fact]
    public async Task HtmlPdf_DirectRendererAsync_AppliesExternalStylesheetPageRules() {
        const string html = "<link rel='stylesheet' href='https://assets.example.test/print.css'><p>ExternalCssPdfMarker</p>";
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
        };
        options.ResourceResolver = (request, cancellationToken) =>
            Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                System.Text.Encoding.UTF8.GetBytes("@page { size:4in 3in; margin:12px; } p { color:#123456; }"),
                "text/css"));

        PdfCore.PdfDocumentConversionResult result = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResultAsync(options);
        byte[] pdf = result.ToBytes();
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
        (double width, double height) = read.Pages[0].GetPageSize();

        Assert.Equal(288D, width, 2);
        Assert.Equal(216D, height, 2);
        Assert.Contains("ExternalCssPdfMarker", read.ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_ExposesSharedRenderResourcePolicy() {
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();
        options.ResourceTimeout = TimeSpan.FromSeconds(5D);
        options.MaxConcurrentResourceLoads = 3;
        options.MaxResourceBytes = 1024L;
        options.MaxTotalResourceBytes = 4096L;
        options.MaxResourceCount = 12;
        options.MaxResourceRequests = 24;
        options.MaxStylesheetImportDepth = 4;
        options.UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        options.ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(null);

        HtmlPdfResourcePolicySummary summary = options.GetResourcePolicySummary();

        Assert.True(summary.HasResourceResolver);
        Assert.True(summary.AllowSystemFontEmbedding);
        Assert.False(summary.AllowDocumentFontEmbedding);
        Assert.False(summary.AllowLocalFileAccess);
        Assert.False(summary.AllowRemoteResourceResolution);
        Assert.True(summary.AllowDataUris);
        Assert.True(summary.AllowEmbeddedPackageResources);
        Assert.Equal(TimeSpan.FromSeconds(5D), summary.ResourceTimeout);
        Assert.Equal(3, summary.MaxConcurrentResourceLoads);
        Assert.Equal(1024L, summary.MaxResourceBytes);
        Assert.Equal(4096L, summary.MaxTotalResourceBytes);
        Assert.Equal(12, summary.MaxResourceCount);
        Assert.Equal(24, summary.MaxResourceRequests);
        Assert.Equal(4, summary.MaxStylesheetImportDepth);
        Assert.Contains("https", summary.AllowedUrlSchemes);
    }

    [Fact]
    public void HtmlPdf_DocumentFontPolicyPreventsCssFamiliesFromLocatingHostFonts() {
        string? installedFamily = new[] { "Arial", "Calibri", "Liberation Sans", "DejaVu Sans" }
            .FirstOrDefault(candidate => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
        if (installedFamily == null) return;

        var options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateDefault()
        };

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument
            .Parse("<p style=\"font-family:'" + installedFamily + "'\">Document font boundary</p>")
            .ToPdfDocumentResult(options);

        Assert.True(options.GetResourcePolicySummary().AllowSystemFontEmbedding);
        Assert.False(options.GetResourcePolicySummary().AllowDocumentFontEmbedding);
        Assert.Empty(result.Value.Options.EmbeddedFonts);
        Assert.Contains("Document font boundary", PdfCore.PdfReadDocument.Open(result.Value.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_TrustedResourcePolicyDoesNotRelaxCallerHyperlinkPolicy() {
        const string webUri = "https://example.test/report";
        const string fileUri = "file:///secret/report.pdf";
        const string dataUri = "data:text/plain,private";
        var options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            UrlPolicy = HtmlUrlPolicy.CreateHyperlinkProfile()
        };

        byte[] pdf = HtmlConversionDocument.Parse($"""
            <p><a href="{webUri}">Web report</a></p>
            <p><a href="{fileUri}">Local report</a></p>
            <p><a href="{dataUri}">Inline report</a></p>
            """).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Contains(webUri, info.LinkUris);
        Assert.DoesNotContain(fileUri, info.LinkUris);
        Assert.DoesNotContain(dataUri, info.LinkUris);
    }

    [Fact]
    public void HtmlPdf_DefaultHyperlinkPolicyEmitsOnlyCommonWebAndMailSchemes() {
        const string webUri = "https://example.test/report";
        const string mailUri = "mailto:reports@example.test";
        const string ftpUri = "ftp://example.test/report";
        const string cidUri = "cid:report";
        const string customUri = "officeimo:report";

        byte[] pdf = HtmlConversionDocument.Parse($"""
            <a href="{webUri}">Web</a>
            <a href="{mailUri}">Mail</a>
            <a href="{ftpUri}">FTP</a>
            <a href="{cidUri}">CID</a>
            <a href="{customUri}">Custom</a>
            """).ToPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Contains(webUri, info.LinkUris);
        Assert.Contains(mailUri, info.LinkUris);
        Assert.DoesNotContain(ftpUri, info.LinkUris);
        Assert.DoesNotContain(cidUri, info.LinkUris);
        Assert.DoesNotContain(customUri, info.LinkUris);
    }

    [Fact]
    public void HtmlPdf_GenericRenderOptionsAdoptPdfHyperlinkDefaults() {
        const string webUri = "https://example.test/report";
        const string ftpUri = "ftp://example.test/report";
        const string customUri = "officeimo:report";
        var sharedOptions = new HtmlRenderOptions {
            ViewportWidth = 720D
        };
        var pdfOptions = new HtmlPdfSaveOptions(sharedOptions);

        byte[] pdf = HtmlConversionDocument.Parse($"""
            <a href="{webUri}">Web</a>
            <a href="{ftpUri}">FTP</a>
            <a href="{customUri}">Custom</a>
            """).ToPdf(pdfOptions);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(720D, pdfOptions.ViewportWidth);
        Assert.Contains(webUri, info.LinkUris);
        Assert.DoesNotContain(ftpUri, info.LinkUris);
        Assert.DoesNotContain(customUri, info.LinkUris);
    }

    [Fact]
    public async Task HtmlPdf_WebOnlyResourcePolicyExpandsOnlyPermittedDataAndFileSchemes() {
        byte[] dataImage = PdfPngTestImages.CreateRgbPng(8, 5);
        byte[] fileImage = PdfPngTestImages.CreateRgbPng(5, 8);
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(dataImage);
        string fileUri = new Uri(Path.Combine(Path.GetTempPath(), "officeimo-web-only-resource.png")).AbsoluteUri;
        int fileResolverCalls = 0;
        var options = new HtmlPdfSaveOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            ResourceResolver = (request, cancellationToken) => {
                if (!request.Uri.IsFile) return Task.FromResult<HtmlResolvedResource?>(null);
                fileResolverCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(fileImage, "image/png"));
            }
        };

        var documentOptions = new HtmlConversionDocumentOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceUrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile()
        };
        PdfCore.PdfDocumentConversionResult result = await HtmlConversionDocument.Parse($"""
            <a href="{dataUri}">Blocked data link</a>
            <a href="{fileUri}">Blocked file link</a>
            <img src="{dataUri}" width="40" height="25" alt="data image">
            <img src="{fileUri}" width="25" height="40" alt="file image">
            """, documentOptions).ToPdfDocumentResultAsync(options);
        byte[] pdf = result.ToBytes();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(1, fileResolverCalls);
        Assert.Equal(2, PdfCore.PdfImageExtractor.ExtractImages(pdf).Count(image => image.IsImageFile && image.MimeType == "image/png"));
        Assert.DoesNotContain(dataUri, info.LinkUris);
        Assert.DoesNotContain(fileUri, info.LinkUris);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public async Task HtmlPdf_PortableResourcePolicyDoesNotInvokeRemoteResolver() {
        int calls = 0;
        var options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            ResourceResolver = (request, cancellationToken) => {
                calls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes("p { color: red; }"),
                    "text/css"));
            }
        };

        PdfCore.PdfDocumentConversionResult result = await HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'><p>Portable</p>")
            .ToPdfDocumentResultAsync(options);

        Assert.Equal(0, calls);
        Assert.Contains(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public async Task HtmlPdf_TrustedHostResourcePolicyInvokesRemoteResolver() {
        int calls = 0;
        var options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            ResourceResolver = (request, cancellationToken) => {
                calls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes("p { color: red; }"),
                    "text/css"));
            }
        };

        PdfCore.PdfDocumentConversionResult result = await HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'><p>Trusted</p>")
            .ToPdfDocumentResultAsync(options);

        Assert.Equal(1, calls);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending);
    }

    [Fact]
    public async Task MhtmlPdf_DefaultPolicyResolvesEmbeddedCidImageThroughDirectLifecycle() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        var archive = new MhtmlDocument(
            "<h1>MhtmlPdfMarker</h1><img src='cid:logo@example.test' width='40' height='25' alt='embedded logo'>",
            new[] { new MhtmlResource(imageBytes, "image/png", contentId: "logo@example.test", fileName: "logo.png") });

        PdfCore.PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync();
        byte[] pdf = result.ToBytes();

        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText().Replace("\r", string.Empty).Replace("\n", string.Empty);
        Assert.Contains("MhtmlPdfMarker", text, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public void MhtmlPdf_SynchronousLifecycleResolvesEmbeddedCidAndContentLocationImages() {
        byte[] cidImage = PdfPngTestImages.CreateRgbPng(8, 5);
        byte[] locationImage = PdfPngTestImages.CreateRgbPng(5, 8);
        var archive = new MhtmlDocument(
            """
            <h1>Synchronous MHTML</h1>
            <a href='cid:logo@example.test'>blocked package link</a>
            <img src='cid:logo@example.test' width='40' height='25' alt='CID logo'>
            <img src='images/chart.png' width='25' height='40' alt='location chart'>
            """,
            new[] {
                new MhtmlResource(cidImage, "image/png", contentId: "logo@example.test", fileName: "logo.png"),
                new MhtmlResource(locationImage, "image/png", contentLocation: "images/chart.png", fileName: "chart.png")
            },
            contentLocation: "https://snapshot.example.test/archive/page.html");

        var options = new HtmlPdfSaveOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };
        PdfCore.PdfDocumentConversionResult result = archive.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();

        Assert.Equal(2, PdfCore.PdfImageExtractor.ExtractImages(pdf).Count(image => image.IsImageFile && image.MimeType == "image/png"));
        Assert.DoesNotContain(PdfCore.PdfInspector.Inspect(pdf).LinkUris, link => link.StartsWith("cid:", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("cid", options.UrlPolicy.AllowedUrlSchemes);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public async Task MhtmlPdf_DefaultPolicyResolvesEmbeddedContentLocationIndependentlyOfUriScheme() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        var archive = new MhtmlDocument(
            "<h1>MhtmlLocationMarker</h1><img src='images/logo.png' width='40' height='25' alt='embedded location logo'>",
            new[] { new MhtmlResource(imageBytes, "image/png", contentLocation: "images/logo.png", fileName: "logo.png") },
            contentLocation: "https://snapshot.example.test/archive/page.html");

        PdfCore.PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync();
        byte[] pdf = result.ToBytes();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public async Task MhtmlPdf_DefaultPolicyResolvesEmbeddedContentLocationFromFileBackedArchive() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        var source = new MhtmlDocument(
            "<a href='local-review.pdf'>blocked local link</a><img src='images/logo.png' width='40' height='25' alt='file-backed embedded logo'>",
            new[] { new MhtmlResource(imageBytes, "image/png", contentLocation: "images/logo.png", fileName: "logo.png") });
        string path = Path.Combine(Path.GetTempPath(), "officeimo-mhtml-pdf-" + Guid.NewGuid().ToString("N") + ".mht");
        try {
            source.Save(path);
            MhtmlDocument archive = MhtmlDocument.Load(path);
            Assert.True(archive.BaseUri.IsFile);

            PdfCore.PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync();
            byte[] pdf = result.ToBytes();

            Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
            Assert.DoesNotContain(PdfCore.PdfInspector.Inspect(pdf).LinkUris, link => link.StartsWith("file:", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task MhtmlPdf_EmbeddedPolicyCannotBeBypassedByTrustedHostUriSchemes() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        var archive = new MhtmlDocument(
            "<img src='https://snapshot.example.test/assets/logo.png' width='40' height='25' alt='blocked embedded logo'>",
            new[] { new MhtmlResource(imageBytes, "image/png", contentLocation: "https://snapshot.example.test/assets/logo.png", fileName: "logo.png") },
            contentLocation: "https://snapshot.example.test/archive/page.html");
        PdfCore.PdfResourcePolicy policy = PdfCore.PdfResourcePolicy.CreateTrustedHost();
        policy.AllowEmbeddedPackageResources = false;

        PdfCore.PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync(new HtmlPdfSaveOptions {
            ResourcePolicy = policy
        });
        byte[] pdf = result.ToBytes();

        Assert.DoesNotContain(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Contains(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public async Task MhtmlPdf_PackageSchemeExpansionDoesNotReachTrustedHostResolver() {
        int hostResolverCalls = 0;
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(8, 5);
        var archive = new MhtmlDocument(
            "<img src='cid:logo' width='40' height='25' alt='embedded logo'><img src='ftp://snapshot.example.test/missing.png' alt='missing'>",
            new[] { new MhtmlResource(imageBytes, "image/png", contentId: "logo", fileName: "logo.png") },
            contentLocation: "ftp://snapshot.example.test/archive/page.html");

        PdfCore.PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync(new HtmlPdfSaveOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            ResourceResolver = (request, cancellationToken) => {
                hostResolverCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(imageBytes, "image/png"));
            }
        });
        byte[] pdf = result.ToBytes();

        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Equal(0, hostResolverCalls);
        Assert.Contains(result.Warnings, warning => warning.Code == "ImageResourceRejectedByPolicy");
    }

    [Fact]
    public void MhtmlPdf_ExposesCompleteDirectLifecycle() {
        MethodInfo[] methods = typeof(MhtmlPdfConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.GetParameters().FirstOrDefault()?.ParameterType == typeof(MhtmlDocument))
            .ToArray();

        Assert.Single(methods, method => method.Name == "ToPdf");
        Assert.Single(methods, method => method.Name == "ToPdfAsync");
        Assert.Single(methods, method => method.Name == "ToPdfDocument");
        Assert.Single(methods, method => method.Name == "ToPdfDocumentAsync");
        Assert.Single(methods, method => method.Name == "ToPdfDocumentResult");
        Assert.Single(methods, method => method.Name == "ToPdfDocumentResultAsync");
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdf"));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfAsync"));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdf"));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdfAsync"));
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
    public void HtmlComputedStyles_DiscardUnknownStylesheetDeclarationsBeforeCascadeStorage() {
        var css = new StringBuilder("<style>.target{");
        for (int index = 0; index < 2_000; index++) css.Append("unknown-").Append(index).Append(":value;");
        css.Append("--brand:#123456;color:var(--brand)}</style><p class='target'>Target</p>");
        var parsed = HtmlDocumentParser.ParseDocument(css.ToString());

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(parsed)[parsed.QuerySelector(".target")!];

        Assert.Equal("#123456", style.GetValue("color"));
        Assert.Equal("#123456", style.GetValue("--brand"));
        Assert.DoesNotContain(style.Properties.Keys, property => property.StartsWith("unknown-", StringComparison.Ordinal));
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlRender_Continuous_BreaksLongTokensAtUnicodeTextElementBoundaries() {
        string composed = "e\u0301";
        string smile = char.ConvertFromUtf32(0x1F600);
        string value = "A" + composed + smile + "B";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 20D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<p style='margin:0;font-size:12px'>" + value + "</p>", options);
        IReadOnlyList<string> segments = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text).ToList();

        Assert.Equal(value, string.Concat(segments));
        Assert.DoesNotContain(segments, segment => segment == "\u0301");
        Assert.DoesNotContain(segments, segment => segment.Length == 1 && char.IsSurrogate(segment[0]));
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
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(20D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<OfficeImageExportResult> images = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options);

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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText text = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), visual => visual.Text.Contains("Page", StringComparison.Ordinal));

        Assert.Equal(480D, page.Width, 3);
        Assert.Equal(288D, page.Height, 3);
        Assert.InRange(text.X, 23.9D, 24.1D);

        options.HonorCssPageRules = false;
        HtmlRenderDocument ignored = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
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

        HtmlRenderPage page = Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options).Pages);

        Assert.Equal(480D, page.Width, 3);
        Assert.Equal(288D, page.Height, 3);
    }

    [Fact]
    public void HtmlRender_Paged_PageRulesUseConfiguredMediaFeatures() {
        string html = "<style>@media (prefers-color-scheme:dark) { @page { size: 5in 3in; } } @media (prefers-color-scheme:light) { @page { size: 2in 2in; } }</style><p>Dark page</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(12D),
            MediaFeatures = new HtmlRenderMediaFeatures {
                PreferredColorScheme = HtmlPreferredColorScheme.Dark
            }
        };

        HtmlRenderPage page = Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options).Pages);

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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), renderOptions);
        string renderedText = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text));

        Assert.True(rendered.Pages.Count >= 3);
        Assert.Contains("word000", renderedText, StringComparison.Ordinal);
        Assert.Contains("word089", renderedText, StringComparison.Ordinal);
        Assert.Contains("Row00", renderedText, StringComparison.Ordinal);
        Assert.Contains("Row17", renderedText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == "HtmlRenderBlockExceedsPage");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(renderOptions);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("word089", pdfText, StringComparison.Ordinal);
        Assert.Contains("Row17", pdfText, StringComparison.Ordinal);
        Assert.Equal(rendered.Pages.Count, PdfCore.PdfInspector.Inspect(pdf).PageCount);
    }

    [Fact]
    public void HtmlRender_Paged_EnforcesWidowsAndOrphansThroughNestedBlocks() {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(16D)
        };
        int selectedWordCount = 0;
        for (int wordCount = 20; wordCount <= 100; wordCount++) {
            string candidateWords = string.Join(" ", Enumerable.Range(0, wordCount).Select(index => "word" + index.ToString("D3")));
            HtmlRenderDocument baseline = HtmlRenderTestDriver.Render("<div><p style='margin:0;orphans:1;widows:1'>" + candidateWords + "</p></div>", options);
            int finalPageLines = CountRenderedTextLines(baseline.Pages[baseline.Pages.Count - 1]);
            if (baseline.Pages.Count > 1 && finalPageLines > 0 && finalPageLines < 4) {
                selectedWordCount = wordCount;
                break;
            }
        }

        Assert.True(selectedWordCount > 0, "The deterministic text corpus should expose a short final fragment without widow protection.");
        string words = string.Join(" ", Enumerable.Range(0, selectedWordCount).Select(index => "word" + index.ToString("D3")));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<div><p style='margin:0;orphans:4;widows:4'>" + words + "</p></div>", options);

        Assert.True(rendered.Pages.Count > 1);
        Assert.All(rendered.Pages, page => Assert.True(CountRenderedTextLines(page) >= 4));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlRender_Paged_RepeatsLeadingTableHeaderRowsInImagesAndSearchablePdf() {
        string rows = string.Join(string.Empty, Enumerable.Range(0, 18).Select(index => "<tr><td>Row" + index.ToString("D2") + "</td></tr>"));
        string html = "<div style='padding:2px'><table><thead><tr><th>HeaderMarker</th></tr></thead><tbody>" + rows + "</tbody></table></div>";
        var renderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), renderOptions);

        Assert.True(rendered.Pages.Count >= 3);
        Assert.All(rendered.Pages, page =>
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "HeaderMarker"));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed);
        IReadOnlyList<OfficeImageExportResult> images = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Png, renderOptions);
        Assert.Equal(rendered.Pages.Count, images.Count);
        Assert.All(images, image => Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, image.Bytes.Take(8)));

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(renderOptions);
        string pdfText = PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
        int repeatedHeaderCount = pdfText.Split(new[] { "HeaderMarker" }, StringSplitOptions.None).Length - 1;
        Assert.Equal(rendered.Pages.Count, repeatedHeaderCount);
    }

    [Fact]
    public void HtmlRender_Paged_RepeatsTableFooterRowsWithoutDuplicatingSourceRows() {
        string rows = string.Join(string.Empty, Enumerable.Range(0, 18).Select(index => "<tr><td>Row" + index.ToString("D2") + "</td></tr>"));
        string html = "<div style='padding:2px'><table><thead><tr><th>HeaderMarker</th></tr></thead><tfoot><tr><td>FooterMarker</td></tr></tfoot><tbody>" + rows + "</tbody></table></div>";
        var renderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), renderOptions);
        string renderedText = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text));

        Assert.True(rendered.Pages.Count >= 3);
        Assert.All(rendered.Pages, page => {
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "HeaderMarker");
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "FooterMarker");
        });
        foreach (int index in Enumerable.Range(0, 18)) {
            string marker = "Row" + index.ToString("D2");
            Assert.Equal(1, renderedText.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        IReadOnlyList<OfficeImageExportResult> images = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Png, renderOptions);
        Assert.Equal(rendered.Pages.Count, images.Count);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(renderOptions);
        string pdfText = PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
        int repeatedFooterCount = pdfText.Split(new[] { "FooterMarker" }, StringSplitOptions.None).Length - 1;
        Assert.Equal(rendered.Pages.Count, repeatedFooterCount);

        HtmlRenderDocument continuous = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous });
        string continuousText = string.Join(" ", continuous.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text));
        Assert.Equal(1, continuousText.Split(new[] { "FooterMarker" }, StringSplitOptions.None).Length - 1);
        Assert.True(continuousText.IndexOf("Row17", StringComparison.Ordinal) < continuousText.IndexOf("FooterMarker", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRender_Paged_LaysOutRowSpansAndKeepsSpanGroupsTogether() {
        string groups = string.Join(string.Empty, Enumerable.Range(0, 10).Select(index =>
            "<tr><td id='span" + index + "' rowspan='2'>Group" + index.ToString("D2") + "</td><td id='regular" + index + "'>Row" + index.ToString("D2") + "A</td></tr>"
            + "<tr><td>Row" + index.ToString("D2") + "B</td></tr>"));
        string html = "<div><table><thead><tr><th>HeaderMarker</th><th>Value</th></tr></thead><tbody>" + groups
            + "</tbody><tbody><tr><td id='zero' rowspan='0'>ZeroMarker</td><td>ZeroA</td></tr><tr><td>ZeroB</td></tr></tbody>"
            + "<tfoot><tr><td>FooterMarker</td><td>End</td></tr></tfoot></table></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<HtmlRenderVisual> visuals = rendered.Pages.SelectMany(page => page.Visuals).ToList();
        string renderedText = string.Join(" ", visuals.OfType<HtmlRenderText>().Select(text => text.Text));

        Assert.True(rendered.Pages.Count >= 3);
        foreach (int index in Enumerable.Range(0, 10)) {
            string marker = "Group" + index.ToString("D2");
            Assert.Equal(1, renderedText.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }

        HtmlRenderShape spanShape = Assert.Single(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#span0");
        HtmlRenderShape regularShape = Assert.Single(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#regular0");
        HtmlRenderShape zeroShape = Assert.Single(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#zero");
        Assert.True(spanShape.Height > regularShape.Height);
        Assert.True(zeroShape.Height > regularShape.Height);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        string pdfText = PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
        Assert.Contains("Group00", pdfText, StringComparison.Ordinal);
        Assert.Contains("Group09", pdfText, StringComparison.Ordinal);
        Assert.Contains("ZeroMarker", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_Paged_RendersFirstLeftRightMarginContentAcrossSvgAndPdf() {
        string words = string.Join(" ", Enumerable.Range(0, 120).Select(index => "word" + index.ToString("D3")));
        string html = """
            <style>
              @page {
                size: 3in 2in;
                margin: 0.3in;
                @top-center { content: "Page " counter(page) " of " counter(pages); color:#224466; font-size:10px; }
                @bottom-right { content: "GenericFooter"; }
              }
              @page :first { @top-center { content: "FirstPage"; font-weight:bold; } }
              @page :left { @bottom-left { content: "L" counter(page); } }
              @page :right { @bottom-right { content: "R" counter(page); } }
            </style>
            <div><p style="margin:0">WORDS</p></div>
            """.Replace("WORDS", words);
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.True(rendered.Pages.Count >= 3);
        Assert.Equal(288D, rendered.Pages[0].Width, 3);
        Assert.Equal(192D, rendered.Pages[0].Height, 3);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.SemanticRole == "page-margin" && text.Text == "FirstPage");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Page 2 of " + rendered.Pages.Count);
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "L2");
        Assert.Contains(rendered.Pages[2].Visuals.OfType<HtmlRenderText>(), text => text.Text == "R3");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageSelectorPending);

        IReadOnlyList<OfficeImageExportResult> svgPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options);
        Assert.Contains("FirstPage", Encoding.UTF8.GetString(svgPages[0].Bytes), StringComparison.Ordinal);
        Assert.Contains("L2", Encoding.UTF8.GetString(svgPages[1].Bytes), StringComparison.Ordinal);
        Assert.Contains("R3", Encoding.UTF8.GetString(svgPages[2].Bytes), StringComparison.Ordinal);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf, new PdfCore.PdfReadOptions { IncludeArtifactText = true }).ExtractText();
        Assert.Equal(rendered.Pages.Count, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Contains("FirstPage", pdfText, StringComparison.Ordinal);
        Assert.Contains("Page 2 of " + rendered.Pages.Count, pdfText, StringComparison.Ordinal);
        Assert.Contains("L2", pdfText, StringComparison.Ordinal);
        Assert.Contains("R3", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_Paged_AppliesFirstPageGeometryAndReflowsTheBody() {
        string html = "<style>@page { size:3in 2in; margin:0.25in; } @page :first { size:2in 2in; margin:0.5in; @top-left { content:\"First\"; } }</style><p>Body</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(192D, page.Width, 3);
        Assert.Equal(192D, page.Height, 3);
        Assert.Equal(48D, page.Margins.Left, 3);
        Assert.Equal(48D, page.Margins.Top, 3);
        HtmlRenderText body = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Body", StringComparison.Ordinal));
        Assert.Equal(48D, body.X, 3);
        Assert.True(body.Width <= 96D + 0.0001D);
        Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "First");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageSelectorPending);
    }

    [Fact]
    public void HtmlRender_Paged_CascadesRelativePageSizeBeforeApplyingItToTheCallerDefault() {
        const string html = "<style>@page { size:landscape !important } @page { size:A4 }</style><p>Body</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = OfficePageSizes.Letter
        };

        HtmlRenderPage page = Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options).Pages);

        Assert.Equal(OfficePageSizes.Letter.HeightInches * HtmlRenderOptions.CssPixelsPerInch, page.Width, 3);
        Assert.Equal(OfficePageSizes.Letter.WidthInches * HtmlRenderOptions.CssPixelsPerInch, page.Height, 3);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsTopLevelBlocksWhenTheNextPageMasterChangesWidth() {
        const string html = """
            <style>
              @page { size:3in 2in; margin:0.25in; }
              @page :first { size:2in 2in; margin:0.5in; }
              p { margin:0; font-size:16px; line-height:20px; }
            </style>
            <div style="height:80px;margin:0;background:#eeeeee"></div>
            <p>Second page wider text</p>
            """;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(192D, rendered.Pages[0].Width, 3);
        Assert.Equal(288D, rendered.Pages[1].Width, 3);
        Assert.Equal(24D, rendered.Pages[1].Margins.Left, 3);
        HtmlRenderText text = Assert.Single(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), item => item.Text.Contains("Second page", StringComparison.Ordinal));
        Assert.Equal(24D, text.X, 3);
        Assert.True(text.Width <= 240D + 0.0001D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsAnInlineContinuationFromLogicalSourceProgress() {
        string expected = string.Join(" ", Enumerable.Range(0, 20).Select(index => "word" + index.ToString("D2")));
        string html = """
            <style>
              @page { size:3in 2in; margin:0.25in; }
              @page :first { size:2in 2in; margin:0.5in; }
              p { margin:0; font-size:16px; line-height:20px; orphans:2; widows:2; }
            </style>
            <p>WORDS</p>
            """.Replace("WORDS", expected);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(192D, rendered.Pages[0].Width, 3);
        Assert.Equal(288D, rendered.Pages[1].Width, 3);
        IReadOnlyList<HtmlRenderText> bodyText = rendered.Pages
            .SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole != "page-margin")
            .ToList();
        string actual = string.Join(" ", bodyText.Select(text => text.Text))
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Aggregate(string.Empty, (current, word) => current.Length == 0 ? word : current + " " + word);
        Assert.Equal(expected, actual);
        Assert.InRange(rendered.Pages[1].Visuals.OfType<HtmlRenderText>().Select(text => text.Y).Distinct().Count(), 1, 6);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsInlineContinuationsWhenOnlyThePageHeightChanges() {
        string words = string.Join(" ", Enumerable.Range(0, 80).Select(index => "height" + index.ToString("D2")));
        string html = """
            <style>
              @page { size:200px 200px; margin:20px; }
              @page :first { size:200px 120px; margin:20px; }
              p { margin:0; font-size:10vh; line-height:1; orphans:2; widows:2; }
            </style>
            <p>WORDS</p>
            """.Replace("WORDS", words);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.True(rendered.Pages.Count >= 2);
        HtmlRenderText firstPageText = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Take(1));
        HtmlRenderText secondPageText = Assert.Single(
            rendered.Pages[1].Visuals.OfType<HtmlRenderText>().Take(1));
        Assert.Equal(12D, firstPageText.Font.Size, 3);
        Assert.Equal(20D, secondPageText.Font.Size, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsInlineContinuationsAcrossAlternatingLeftAndRightMasters() {
        string expected = string.Join(" ", Enumerable.Range(0, 50).Select(index => "item" + index.ToString("D2")));
        string html = """
            <style>
              @page { size:3in 1.5in; margin:0.25in; }
              @page :left { size:2in 1.5in; }
              @page :right { size:3in 1.5in; }
              p { margin:0; font-size:16px; line-height:20px; orphans:2; widows:2; }
            </style>
            <p><span>WORDS</span></p>
            """.Replace("WORDS", expected);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.True(rendered.Pages.Count >= 3);
        Assert.All(rendered.Pages.Where(page => page.PageNumber % 2 == 0), page => Assert.Equal(192D, page.Width, 3));
        Assert.All(rendered.Pages.Where(page => page.PageNumber % 2 != 0), page => Assert.Equal(288D, page.Width, 3));
        string actual = string.Join(" ", rendered.Pages
            .SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>()
            .Select(text => text.Text))
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Aggregate(string.Empty, (current, word) => current.Length == 0 ? word : current + " " + word);
        Assert.Equal(expected, actual);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsNestedBlockContinuationsAcrossAlternatingPageMasters() {
        string words = string.Join(" ", Enumerable.Range(0, 50).Select(index => "nested" + index.ToString("D2")));
        string expected = "Intro " + words + " Outro";
        string html = """
            <style>
              @page { size:3in 1.5in; margin:0.25in; }
              @page :left { size:2in 1.5in; }
              @page :right { size:3in 1.5in; }
              section, div, p { margin:0; padding:0; }
              p { font-size:16px; line-height:20px; orphans:2; widows:2; }
            </style>
            <section id="wrapper">
              <div id="intro">Intro</div>
              <div id="nested"><p><span>WORDS</span></p></div>
              <div id="outro">Outro</div>
            </section>
            """.Replace("WORDS", words);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.True(rendered.Pages.Count >= 3);
        string actual = string.Join(" ", rendered.Pages
            .SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>()
            .Select(text => text.Text))
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Aggregate(string.Empty, (current, word) => current.Length == 0 ? word : current + " " + word);
        Assert.Equal(expected, actual);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsAtNestedChildBoundaryWhenFinalChildNeedsNoFurtherBreak() {
        string html = """
            <style>
              @page { size:3in 1in; margin:0; }
              @page :left { size:2in 1in; }
              section, div { margin:0; padding:0; }
              #first { height:96px; }
              #last { font-size:16px; line-height:20px; }
            </style>
            <section><div id="first">First</div><div id="last">Last</div></section>
            """;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(288D, rendered.Pages[0].Width, 3);
        Assert.Equal(192D, rendered.Pages[1].Width, 3);
        Assert.Equal(
            new[] { "First", "Last" },
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text).ToArray());
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsTableRowsAcrossAlternatingPageMasters() {
        string rows = string.Join(string.Empty, Enumerable.Range(0, 18).Select(index =>
            "<tr><td>Row" + index.ToString("D2") + "</td><td>Value" + index.ToString("D2") + "</td></tr>"));
        string html = """
            <style>
              @page { size:3in 1.75in; margin:0.2in; }
              @page :left { size:2in 1.75in; }
              table { width:100%; border-spacing:0; }
              th, td { padding:2px; font-size:12px; line-height:14px; }
            </style>
            <table><thead><tr><th>Header</th><th>Value</th></tr></thead><tbody>ROWS</tbody></table>
            """.Replace("ROWS", rows);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.True(rendered.Pages.Count >= 3);
        Assert.All(rendered.Pages.Where(page => page.PageNumber % 2 == 0), page => Assert.Equal(192D, page.Width, 3));
        Assert.All(rendered.Pages.Where(page => page.PageNumber % 2 != 0), page => Assert.Equal(288D, page.Width, 3));
        Assert.All(rendered.Pages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header"));
        string renderedText = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text));
        foreach (int index in Enumerable.Range(0, 18)) {
            string marker = "Row" + index.ToString("D2");
            Assert.Equal(1, renderedText.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlRender_Paged_ReflowsWrappedFlexLinesAcrossAlternatingPageMasters() {
        string items = string.Join(string.Empty, Enumerable.Range(0, 18).Select(index =>
            "<div class='item'>Flex" + index.ToString("D2") + "</div>"));
        string html = """
            <style>
              @page { size:3in 1.5in; margin:0.2in; }
              @page :left { size:2in 1.5in; }
              .flex { display:flex; flex-wrap:wrap; gap:0; width:100%; }
              .item { box-sizing:border-box; width:80px; height:30px; margin:0; padding:2px; font-size:12px; line-height:14px; }
            </style>
            <div class="flex">ITEMS</div>
            """.Replace("ITEMS", items);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.True(rendered.Pages.Count >= 3);
        Assert.All(rendered.Pages.Where(page => page.PageNumber % 2 == 0), page => Assert.Equal(192D, page.Width, 3));
        Assert.All(rendered.Pages.Where(page => page.PageNumber % 2 != 0), page => Assert.Equal(288D, page.Width, 3));
        string renderedText = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(text => text.Text));
        foreach (int index in Enumerable.Range(0, 18)) {
            string marker = "Flex" + index.ToString("D2");
            Assert.Equal(1, renderedText.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlRender_Paged_DiagnosesComplexPageSelectorsUnknownMarginPositionsAndGeneratedContent() {
        string html = "<style>@page invoice:first:right { @top-left { content:\"Complex\"; } } @page { @left-middle { content:\"Side\"; } @unknown-zone { content:\"Unknown\"; } @top-left { content:attr(title); } }</style><p>Body</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(16D)
        });

        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageSelectorPending);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageMarginContentUnsupported);
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "Side");
    }

    [Fact]
    public void HtmlRender_Paged_ChargesRepeatedMarginTextToTheLayoutBudgetBeforeAllocation() {
        string html = "<style>@page { @top-left { content:\"" + new string('A', 500) + "\"; } }</style><p>Body</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            MaxLayoutOperations = 100
        };

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }

    [Fact]
    public void HtmlRender_Paged_AppliesNamedPageMastersAndNamedPseudoPages() {
        string words = string.Join(" ", Enumerable.Range(0, 150).Select(index => "word" + index.ToString("D3")));
        string html = """
            <style>
              @page { size:3in 2in; margin:0.3in; @top-left { content:"Generic"; } }
              @page invoice { @top-left { content:"Invoice"; } }
              @page invoice:left { @bottom-left { content:"IL"; } }
              @page report { @top-left { content:"Report"; } }
            </style>
            <section style="page:invoice"><p style="margin:0">WORDS</p></section>
            <section style="page:report"><p style="margin:0">ReportBody</p></section>
            """.Replace("WORDS", words);
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<HtmlRenderPage> invoicePages = rendered.Pages.Where(page => page.PageName == "invoice").ToList();
        HtmlRenderPage reportPage = Assert.Single(rendered.Pages, page => page.PageName == "report");

        Assert.True(invoicePages.Count >= 2);
        Assert.All(invoicePages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Invoice"));
        Assert.Contains(invoicePages.Where(page => page.PageNumber % 2 == 0).SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "IL");
        Assert.Contains(reportPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Report");
        Assert.Contains(reportPage.Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("ReportBody", StringComparison.Ordinal));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageSelectorPending);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        string pdfText = PdfCore.PdfReadDocument.Open(
            OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions),
            new PdfCore.PdfReadOptions { IncludeArtifactText = true }).ExtractText();
        Assert.Contains("Invoice", pdfText, StringComparison.Ordinal);
        Assert.Contains("IL", pdfText, StringComparison.Ordinal);
        Assert.Contains("Report", pdfText, StringComparison.Ordinal);
        Assert.Contains("ReportBody", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPagedMedia_ReflowsNamedPageGeometryAfterBreakAfterCreatedAnEmptyPage() {
        const string html = """
            <style>
              @page { size:300px 200px; margin:10px; }
              @page report { size:200px 120px; margin:20px; }
            </style>
            <div style="break-after:page">First</div>
            <div style="page:report">Second</div>
            """;
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(300D, rendered.Pages[0].Width, 3);
        Assert.Equal("report", rendered.Pages[1].PageName);
        Assert.Equal(200D, rendered.Pages[1].Width, 3);
        Assert.Equal(120D, rendered.Pages[1].Height, 3);
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Second", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRender_Paged_RendersCornerAndSideMarginBoxesAcrossSvgAndPdf() {
        string html = """
            <style>
              @page {
                size: 3in 3in;
                margin: 0.4in;
                @top-left-corner { content:"TLC"; }
                @top-right-corner { content:"TRC"; }
                @left-top { content:"LT"; }
                @left-middle { content:"LM"; }
                @left-bottom { content:"LB"; }
                @right-top { content:"RT"; }
                @right-middle { content:"RM"; }
                @right-bottom { content:"RB"; }
                @bottom-left-corner { content:"BLC"; }
                @bottom-right-corner { content:"BRC"; }
              }
            </style>
            <p>Body</p>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };
        string[] markers = { "TLC", "TRC", "LT", "LM", "LB", "RT", "RM", "RB", "BLC", "BRC" };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        IReadOnlyList<string> visualText = page.Visuals.OfType<HtmlRenderText>().Select(text => text.Text).ToList();
        foreach (string marker in markers) Assert.Contains(marker, visualText);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported);

        string svg = Encoding.UTF8.GetString(Assert.Single(HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options)).Bytes);
        foreach (string marker in markers) Assert.Contains(marker, svg, StringComparison.Ordinal);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        string pdfText = PdfCore.PdfReadDocument.Open(
            OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions),
            new PdfCore.PdfReadOptions { IncludeArtifactText = true }).ExtractText();
        foreach (string marker in markers) Assert.Contains(marker, pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_Paged_RunningElementsUsePageSelectionAndRemainPdfArtifacts() {
        const string html = """
            <style>
              @page {
                size: 320px 180px;
                margin: 32px;
                @top-left { content: element(chapter, start); text-align:left; }
                @top-right { content: element(chapter, first); text-align:right; }
              }
              .running { position: running(chapter); margin:0; padding:2px; border-bottom:1px solid #225588; color:#225588; }
              .page { margin:0; height:80px; }
            </style>
            <h1 class="running">Chapter Alpha</h1>
            <div class="page" style="break-after:page">First body</div>
            <h1 class="running">Chapter Beta</h1>
            <div class="page">Second body</div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        string[] firstPageText = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
        string[] secondPageText = EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
        Assert.Contains("ChapterAlpha", string.Concat(firstPageText), StringComparison.Ordinal);
        Assert.Contains("ChapterAlpha", string.Concat(secondPageText), StringComparison.Ordinal);
        Assert.Contains("ChapterBeta", string.Concat(secondPageText), StringComparison.Ordinal);
        Assert.DoesNotContain("ChapterBeta", string.Concat(firstPageText), StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter Alpha", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter Beta", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("First body", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("Second body", rendered.Text, StringComparison.Ordinal);
        Assert.Empty(rendered.Headings);
        Assert.All(
            rendered.Pages.SelectMany(page => page.Scene).OfType<HtmlRenderSemanticGroup>().Where(group => group.Source?.Contains("element(chapter)", StringComparison.Ordinal) == true),
            group => {
                Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, group.Role);
                Assert.All(
                    EnumerateRenderVisuals(group.Visuals).OfType<HtmlRenderText>(),
                    text => Assert.InRange(text.X, group.X - 0.001D, group.X + group.Width - 0.001D));
            });
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(position: running(chapter))"));
        Assert.Empty(rendered.Diagnostics);

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string raw = Encoding.ASCII.GetString(pdf);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).Outlines);
        Assert.Contains("/Artifact BMC", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter Alpha", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter Beta", pdfText, StringComparison.Ordinal);
        Assert.Contains("First body", pdfText, StringComparison.Ordinal);
        Assert.Contains("Second body", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRenderDocument_TextExcludesArtifactsNestedInsideLogicalTextGroups() {
        var font = new OfficeFontInfo("Arial", 10D);
        var logicalText = new HtmlRenderLogicalTextGroup(
            "Before Decorative After",
            0D,
            0D,
            100D,
            12D,
            new HtmlRenderVisual[] {
                new HtmlRenderText("Before ", 0D, 0D, 30D, 12D, font, OfficeColor.Black, OfficeTextAlignment.Left, 12D, 0),
                new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.Artifact,
                    30D,
                    0D,
                    40D,
                    12D,
                    new[] { new HtmlRenderText("Decorative ", 30D, 0D, 40D, 12D, font, OfficeColor.Black, OfficeTextAlignment.Left, 12D, 0) },
                    1,
                    "decorative"),
                new HtmlRenderText("After", 70D, 0D, 30D, 12D, font, OfficeColor.Black, OfficeTextAlignment.Left, 12D, 2)
            },
            0,
            "logical");
        var rendered = new HtmlRenderDocument(
            HtmlRenderMode.Continuous,
            new[] { new HtmlRenderPage(1, 100D, 100D, new[] { logicalText }) },
            new HtmlDiagnosticReport());

        Assert.Equal("Before After", rendered.Text);
    }

    [Fact]
    public void HtmlRender_Paged_NestedInlineRunningElementLeavesFlowAndRepeatsAsArtifact() {
        const string html = """
            <style>
              @page { size:320px 180px; margin:32px; @top-center { content:element(chapter); } }
              .running { position:running(chapter); color:#225588; border-bottom:1px solid #225588; }
              .page { margin:0; height:80px; }
            </style>
            <p>Before <span class="running">Nested Chapter</span> After</p>
            <div class="page" style="break-after:page">First body</div>
            <div class="page">Second body</div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.All(rendered.Pages, page =>
            Assert.Contains(
                EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(),
                text => text.Text.Contains("Nested", StringComparison.Ordinal)));
        Assert.DoesNotContain(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Text.Contains("Nested Chapter", StringComparison.Ordinal) && text.SemanticRole != "page-margin");
        Assert.DoesNotContain("Nested Chapter", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("Before", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("After", rendered.Text, StringComparison.Ordinal);
        Assert.All(
            rendered.Pages.SelectMany(page => page.Scene).OfType<HtmlRenderSemanticGroup>().Where(group => group.Source?.Contains("element(chapter)", StringComparison.Ordinal) == true),
            group => Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, group.Role));
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("block")]
    [InlineData("inline")]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlRender_Paged_CapturedRunningElementsPropagateOwnAndNestedAssignments(string layout) {
        const string captured = "<span class='outer'>OuterMarker <span class='nested'>NestedMarker</span></span>";
        string body = layout switch {
            "inline" => "<p>Before " + captured + " After</p>",
            "flex" => "<div style='display:flex'>" + captured + "<span>Body</span></div>",
            "grid" => "<div style='display:grid;grid-template-columns:1fr'>" + captured + "<span>Body</span></div>",
            _ => captured + "<p>Body</p>"
        };
        string html = """
            <style>
              @page {
                size:320px 180px;
                margin:32px;
                @top-left { content:string(title); }
                @top-center { content:element(outer); }
                @top-right { content:element(nested); }
              }
              .outer { position:running(outer); string-set:title 'TitleMarker'; }
              .nested { position:running(nested); }
            </style>
            """ + body;

        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        string marginText = string.Concat(
            EnumerateRenderVisuals(rendered.Pages[0].Scene)
                .OfType<HtmlRenderText>()
                .Select(text => text.Text));

        Assert.Contains("TitleMarker", marginText, StringComparison.Ordinal);
        Assert.Contains("OuterMarker", marginText, StringComparison.Ordinal);
        Assert.Contains("NestedMarker", marginText, StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
        PdfCore.PdfDocumentConversionResult pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html)
            .ToPdfDocumentResult(new HtmlPdfSaveOptions(options));
        Assert.DoesNotContain(pdf.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlRender_Paged_DirectFlexAndGridRunningItemsPreserveCaseAndContainerLayout() {
        const string html = """
            <style>
              @page {
                size:320px 200px;
                margin:32px;
                @top-center { content:element(FlexHeader); }
                @bottom-center { content:element(GridFooter); }
              }
              .flex { display:flex; width:240px; height:32px; }
              .grid { display:grid; grid-template-columns:1fr; width:240px; height:32px; }
              .flex-running { position:running(FlexHeader); }
              .grid-running { position:running(GridFooter); }
            </style>
            <div class="flex"><header class="flex-running">MixedCaseHeader</header><div>FlexBody</div></div>
            <div class="grid"><footer class="grid-running">MixedCaseFooter</footer><div>GridBody</div></div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        string sceneText = string.Concat(EnumerateRenderVisuals(Assert.Single(rendered.Pages).Scene).OfType<HtmlRenderText>().Select(text => text.Text));

        Assert.Contains("MixedCaseHeader", sceneText, StringComparison.Ordinal);
        Assert.Contains("MixedCaseFooter", sceneText, StringComparison.Ordinal);
        Assert.Contains("FlexBody", sceneText, StringComparison.Ordinal);
        Assert.Contains("GridBody", sceneText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code is HtmlRenderDiagnosticCodes.FlexLayoutPending or HtmlRenderDiagnosticCodes.GridLayoutPending);
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlRender_Paged_FlexAndGridRunningElementsUseDomOrderForFirstAndLastSelection(string layout) {
        string html = """
            <style>
              @page {
                size:320px 180px;
                margin:32px;
                @top-left { content:element(chapter, first); }
                @top-right { content:element(chapter, last); }
              }
              .container { display:LAYOUT; width:240px; height:64px; }
              .nested { padding-top:20px; }
              .running { position:running(chapter); }
            </style>
            <div class="container">
              <div class="nested"><span class="running">Earlier nested</span></div>
              <span class="running">Later direct</span>
              <div>Body</div>
            </div>
            """.Replace("LAYOUT", layout);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        IReadOnlyList<HtmlRenderSemanticGroup> marginGroups = Assert.Single(rendered.Pages).Scene
            .OfType<HtmlRenderSemanticGroup>()
            .Where(group => group.Source?.Contains("element(chapter", StringComparison.Ordinal) == true)
            .OrderBy(group => group.X)
            .ToList();

        Assert.Equal(2, marginGroups.Count);
        Assert.Contains("Earliernested", string.Concat(EnumerateRenderVisuals(marginGroups[0].Visuals).OfType<HtmlRenderText>().Select(text => text.Text)), StringComparison.Ordinal);
        Assert.Contains("Later direct", string.Concat(EnumerateRenderVisuals(marginGroups[1].Visuals).OfType<HtmlRenderText>().Select(text => text.Text)), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("flex;flex-direction:column")]
    [InlineData("flex;flex-wrap:wrap")]
    [InlineData("grid;grid-template-columns:1fr")]
    public void HtmlRender_Paged_FlexAndGridRunningElementsBecomeAvailableAtTheirFragmentedSourcePosition(string layout) {
        string html = """
            <style>
              @page {
                size:320px 180px;
                margin:24px;
                @top-center { content:element(chapter, last); }
              }
              .container { display:LAYOUT; width:240px; }
              .running { position:running(chapter); }
              .tall { width:240px; height:100px; }
            </style>
            <div class="container">
              <span class="running">Earlier chapter</span>
              <div class="tall">Tall body one</div>
              <div class="tall">Tall body two</div>
              <div class="tall">Tall body three</div>
              <span class="running">Later chapter</span>
            </div>
            """.Replace("LAYOUT", layout);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        IReadOnlyList<string> pageHeaders = rendered.Pages
            .Select(page => string.Concat(page.Scene
                .OfType<HtmlRenderSemanticGroup>()
                .Where(group => group.Source?.Contains("element(chapter", StringComparison.Ordinal) == true)
                .SelectMany(group => EnumerateRenderVisuals(group.Visuals).OfType<HtmlRenderText>())
                .Select(text => text.Text)))
            .ToList();

        Assert.True(pageHeaders.Count > 1);
        Assert.Contains("Earlierchapter", pageHeaders[0], StringComparison.Ordinal);
        Assert.DoesNotContain("Laterchapter", pageHeaders[0], StringComparison.Ordinal);
        Assert.Contains("Laterchapter", pageHeaders[pageHeaders.Count - 1], StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlRender_Paged_MixedRunningElementContentIsDiagnosedAndStrictModeRejectsIt() {
        const string html = """
            <style>
              @page { size:320px 180px; margin:32px; @top-center { content:"Report " element(chapter); } }
              header { position:running(chapter); }
            </style>
            <header>Chapter</header><p>Body</p>
            """;

        HtmlRenderDocument permissive = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlDiagnostic diagnostic = Assert.Single(permissive.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.DoesNotContain(permissive.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.Contains("Report", StringComparison.Ordinal));

        HtmlConversionException exception = Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        }));
        Assert.Contains(exception.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
    }

    [Fact]
    public void HtmlRender_Paged_RightBreakInsertsAStyledBlankLeftPage() {
        string html = """
            <style>
              @page { size:3in 2in; margin:0.25in; }
              @page :left { @top-left { content:"L" counter(page); } }
              @page :right { @top-right { content:"R" counter(page); } }
            </style>
            <p>FirstBody</p>
            <div style="break-before:right">RightBody</div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FirstBody");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "L2" && text.SemanticRole == "page-margin");
        Assert.DoesNotContain(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.SemanticRole != "page-margin");
        Assert.Contains(rendered.Pages[2].Visuals.OfType<HtmlRenderText>(), text => text.Text == "RightBody");
        Assert.Contains(rendered.Pages[2].Visuals.OfType<HtmlRenderText>(), text => text.Text == "R3" && text.SemanticRole == "page-margin");

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf, new PdfCore.PdfReadOptions { IncludeArtifactText = true }).ExtractText();
        Assert.Equal(3, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Contains("FirstBody", pdfText, StringComparison.Ordinal);
        Assert.Contains("L2", pdfText, StringComparison.Ordinal);
        Assert.Contains("RightBody", pdfText, StringComparison.Ordinal);
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
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 320D,
            Margins = HtmlRenderMargins.All(10D)
        };

        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult svg = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options);

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
    public void HtmlRenderedOutputs_AreDeterministicForIdenticalResolvedInput() {
        const string html = "<style>body{margin:0}.card{width:180px;padding:8px;border:2px solid #123456;background:linear-gradient(90deg,#ffffff,#ddeeff)}</style>"
            + "<div class='card'><h2>StableMarker</h2><a href='https://example.test/report'>Report link</a></div>";
        static HtmlRenderOptions ImageOptions() => new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(10D)
        };
        static HtmlPdfSaveOptions PdfOptions() {
            HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();
            options.PageSize = new OfficePageSize(4D, 3D);
            options.Margins = HtmlRenderMargins.All(12D);
            return options;
        }

        byte[] firstPng = HtmlConversionDocument.Parse(html).ToPng(ImageOptions());
        byte[] secondPng = HtmlConversionDocument.Parse(html).ToPng(ImageOptions());
        string firstSvg = HtmlConversionDocument.Parse(html).ToSvg(ImageOptions());
        string secondSvg = HtmlConversionDocument.Parse(html).ToSvg(ImageOptions());
        byte[] firstPdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(PdfOptions());
        byte[] secondPdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(PdfOptions());

        Assert.Equal(firstPng, secondPng);
        Assert.Equal(firstSvg, secondSvg);
        Assert.Equal(firstPdf, secondPdf);
        Assert.Contains("StableMarker", PdfCore.PdfReadDocument.Open(firstPdf).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(HtmlRenderMode.Continuous, 320D, 96D, "Arial")]
    [InlineData(HtmlRenderMode.Continuous, 640D, 144D, "Times New Roman")]
    [InlineData(HtmlRenderMode.Paged, 384D, 192D, "Courier New")]
    public void HtmlRenderingProfiles_StayDeterministicAcrossApprovedSizeDpiAndFontMatrix(
        HtmlRenderMode mode,
        double width,
        double targetDpi,
        string fontFamily) {
        const string html = "<style>body{margin:0}main{display:grid;grid-template-columns:1fr 2fr;gap:8px}h1{font-size:20px}p{margin:0}</style>"
            + "<main><h1>Profile</h1><p>Baseline text 0123456789</p><p dir='rtl'>RTL \u202Babc 123\u202C</p></main>";
        var options = new HtmlRenderOptions {
            Mode = mode,
            ViewportWidth = width,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(12D),
            DefaultFontFamily = fontFamily,
            TargetDpi = targetDpi,
            MaximumRasterPixels = 20_000_000L
        };

        OfficeImageExportResult first = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult second = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(first.Bytes, second.Bytes);
        Assert.Equal(first.Width, second.Width);
        Assert.Equal(first.Height, second.Height);
        Assert.InRange(first.DpiX, targetDpi - 0.1D, targetDpi + 0.1D);
        Assert.InRange(first.DpiY, targetDpi - 0.1D, targetDpi + 0.1D);
        Assert.DoesNotContain(first.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlComputedStyle_DirAttributeParticipatesAsAnOverridablePresentationalHint() {
        const string html = "<!doctype html><html id='root' dir='rtl' style='direction:ltr'><body id='body'><p id='rtl' dir='rtl'><span id='child'>Text</span></p></body></html>";

        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(html);

        Assert.Equal("ltr", styles.Single(pair => pair.Key.Id == "root").Value.GetValue("direction"));
        Assert.Equal("ltr", styles.Single(pair => pair.Key.Id == "body").Value.GetValue("direction"));
        Assert.Equal("rtl", styles.Single(pair => pair.Key.Id == "rtl").Value.GetValue("direction"));
        Assert.Equal("rtl", styles.Single(pair => pair.Key.Id == "child").Value.GetValue("direction"));
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_MapsHeadingsAndParagraphsToTaggedStructure() {
        const string html = "<!doctype html><html lang='pl-PL' dir='rtl'><head><title>Semantic document</title><meta name='author' content='HTML Team'><meta name='description' content='Tagged document proof'><meta name='keywords' content='html, pdf, tagged'><meta name='generator' content='Report Builder'></head><body><main><h1>Semantic <em>heading</em></h1><p>Semantic <strong>paragraph</strong>.</p><h2>Nested detail</h2></main></body></html>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);
        Assert.Equal("Semantic document", rendered.Metadata.Title);
        Assert.Equal("pl-PL", rendered.Metadata.Language);
        Assert.Equal(HtmlRenderTextDirection.RightToLeft, rendered.Metadata.Direction);
        Assert.Equal("HTML Team", rendered.Metadata.Author);
        Assert.Equal("Tagged document proof", rendered.Metadata.Subject);
        Assert.Equal("html, pdf, tagged", rendered.Metadata.Keywords);
        Assert.Equal("Report Builder", rendered.Metadata.Creator);
        Assert.Equal("Semantic document", info.Metadata.Title);
        Assert.Equal("HTML Team", info.Metadata.Author);
        Assert.Equal("Tagged document proof", info.Metadata.Subject);
        Assert.Equal("html, pdf, tagged", info.Metadata.Keywords);
        Assert.Equal("pl-PL", info.CatalogLanguage);
        PdfCore.PdfViewerPreferences viewerPreferences = Assert.IsType<PdfCore.PdfViewerPreferences>(info.ViewerPreferences);
        Assert.Equal("true", viewerPreferences.GetValue("DisplayDocTitle"));
        Assert.Equal("R2L", viewerPreferences.GetValue("Direction"));
        Assert.Collection(
            rendered.Headings,
            heading => {
                Assert.Equal(1, heading.Level);
                Assert.Equal("Semantic heading", heading.Text);
                Assert.Equal(1, heading.PageNumber);
            },
            heading => {
                Assert.Equal(2, heading.Level);
                Assert.Equal("Nested detail", heading.Text);
                Assert.Equal(1, heading.PageNumber);
            });
        Assert.Contains("Document", tagged.StructureTypes);
        Assert.Contains("H1", tagged.StructureTypes);
        Assert.Contains("H2", tagged.StructureTypes);
        Assert.Contains("P", tagged.StructureTypes);
        Assert.Equal(1, tagged.StructureElements.Count(element => element.StructureType == "Sect"));
        Assert.Equal(1, tagged.StructureElements.Count(element => element.StructureType == "H1"));
        Assert.Equal(1, tagged.StructureElements.Count(element => element.StructureType == "H2"));
        Assert.Equal(1, tagged.StructureElements.Count(element => element.StructureType == "P"));
        HtmlRenderSemanticGroup sectionScene = Assert.Single(rendered.Pages[0].Scene.OfType<HtmlRenderSemanticGroup>());
        Assert.Equal(HtmlRenderSemanticGroupRole.Section, sectionScene.Role);
        Assert.Contains(sectionScene.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Heading1);
        Assert.Contains(sectionScene.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Paragraph);
        PdfCore.PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, element => element.StructureType == "Sect");
        Assert.All(
            tagged.StructureElements.Where(element => element.StructureType == "H1" || element.StructureType == "H2" || element.StructureType == "P"),
            element => Assert.Contains(element.ObjectNumber, section.ChildElementObjectNumbers));
        Assert.True(tagged.StructureElements.Count(element => element.StructureType == "Span") >= 5);
        Assert.True(tagged.MarkedContentReferenceCount >= 2);
        PdfCore.PdfOutlineItem outline = Assert.Single(info.Outlines);
        Assert.Equal("Semantic heading", outline.Title);
        Assert.Equal(1, outline.Level);
        Assert.Equal("Nested detail", Assert.Single(outline.Children).Title);
        Assert.Contains("Semantic heading", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_PreservesExplicitUntaggedDocumentMode() {
        var options = new HtmlPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                TaggedStructureMode = PdfCore.PdfTaggedStructureMode.None
            }
        };

        byte[] pdf = HtmlConversionDocument.Parse("<p>Explicit untagged output</p>").ToPdf(options);

        Assert.False(PdfCore.PdfReadDocument.Open(pdf).HasTaggedContent);
    }

    [Fact]
    public void HtmlPdf_CssBookmarkAndTagControlsMapToCanonicalPdfModels() {
        const string html = "<main>"
            + "<h1 style='bookmark-label:\"Public label\";bookmark-state:closed'>Visible title</h1>"
            + "<h2>Child</h2>"
            + "<h2 style='bookmark-level:none'>Suppressed child</h2>"
            + "<div style='bookmark-level:1;bookmark-label:\"Extra entry\";bookmark-state:open'>Extra body</div>"
            + "<h2 style='bookmark-label:\"Empty heading\"'></h2>"
            + "<div style='height:12px;bookmark-level:1;bookmark-label:\"Empty block\"'></div>"
            + "<p>Before<span style='bookmark-level:1;bookmark-label:\"Empty inline\"'></span>After</p>"
            + "<p style='-officeimo-pdf-tag-type:H2'>Promoted semantic paragraph</p>"
            + "<aside style='-officeimo-pdf-tag-type:artifact'>Decorative note</aside>"
            + "</main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions {
                CompressContentStreams = false,
                OutlineExpansionLevel = 64,
                TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers
            }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);
        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Collection(rendered.Headings,
            heading => { Assert.Equal("Public label", heading.Text); Assert.Equal(HtmlRenderBookmarkState.Closed, heading.BookmarkState); },
            heading => Assert.Equal("Child", heading.Text),
            heading => { Assert.Equal("Extra entry", heading.Text); Assert.Equal(HtmlRenderBookmarkState.Open, heading.BookmarkState); },
            heading => Assert.Equal("Empty heading", heading.Text),
            heading => Assert.Equal("Empty block", heading.Text),
            heading => Assert.Equal("Empty inline", heading.Text));
        Assert.Equal("Public label", info.Outlines[0].Title);
        Assert.False(info.Outlines[0].IsExpanded);
        Assert.Equal("Child", Assert.Single(info.Outlines[0].Children).Title);
        Assert.Equal("Extra entry", info.Outlines[1].Title);
        Assert.Equal("Empty heading", Assert.Single(info.Outlines[1].Children).Title);
        Assert.Equal("Empty block", info.Outlines[2].Title);
        Assert.Equal("Empty inline", info.Outlines[3].Title);
        Assert.Contains("/Artifact BMC", raw, StringComparison.Ordinal);
        Assert.Equal(3, tagged.StructureElements.Count(element => element.StructureType == "H2"));
        Assert.DoesNotContain("Decorative note", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(bookmark-level:2)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(bookmark-label:'Label')"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(-officeimo-pdf-tag-type:artifact)"));
    }

    [Theory]
    [InlineData("(bookmark-level:var(--outline-level))")]
    [InlineData("(bookmark-state:var(--outline-state))")]
    [InlineData("(bookmark-label:var(--outline-label))")]
    [InlineData("(clip-path:var(--clip))")]
    [InlineData("(-officeimo-pdf-tag-type:var(--tag))")]
    [InlineData("(bookmark-level:inherit)")]
    [InlineData("(bookmark-state:initial)")]
    [InlineData("(bookmark-label:revert-layer)")]
    public void HtmlSupports_DeferredAndCssWideValuesReachSpecializedProperties(string condition) {
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports(condition));
    }

    [Fact]
    public void HtmlSupports_DeferredValuesDoNotMakeUnknownPropertiesSupported() {
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(not-a-real-property:var(--value))"));
    }

    [Theory]
    [InlineData("(clip-path:var())")]
    [InlineData("(bookmark-level:notvar(--level))")]
    [InlineData("(bookmark-state:var(--state)")]
    [InlineData("(bookmark-label:var(--label, var()))")]
    [InlineData("(bookmark-label:var(--label, [{]}))")]
    public void HtmlSupports_RejectsMalformedDeferredValues(string condition) {
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports(condition));
    }

    [Fact]
    public void HtmlPdf_CssBookmarkAndTagControlsReachSpecializedAndInlineElements() {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(8, 6));
        string html = "<main>"
            + "<img src='data:image/png;base64," + image + "' alt='Chart' style='display:block;width:8px;height:6px;bookmark-level:1;bookmark-label:\"Chart entry\";-officeimo-pdf-tag-type:H2'>"
            + "<p style='width:72px'>Prefix <span style='bookmark-level:2;bookmark-label:\"Inline entry\";-officeimo-pdf-tag-type:H3'>Inline content wraps across lines</span></p>"
            + "<p><input name='field' aria-label='Field' value='Value' style='width:80px;bookmark-level:1;bookmark-label:\"Field entry\";-officeimo-pdf-tag-type:H4'></p>"
            + "<table style='-officeimo-pdf-tag-type:artifact'><tr><td>Decorative table<input name='artifact-field' value='Must not be interactive'></td></tr></table>"
            + "</main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            ResourceUrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile(),
            PdfOptions = new PdfCore.PdfOptions {
                CompressContentStreams = false,
                TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers
            }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        IReadOnlyList<HtmlRenderSemanticGroup> groups = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderSemanticGroup>().ToList();

        Assert.NotEmpty(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderBookmarkAnchor>());
        Assert.Contains(rendered.Headings, heading => heading.Text == "Chart entry" && heading.Level == 1);
        Assert.Contains(rendered.Headings, heading => heading.Text == "Inline entry" && heading.Level == 2);
        Assert.Contains(rendered.Headings, heading => heading.Text == "Field entry" && heading.Level == 1);
        Assert.Contains(groups, group => group.Role == HtmlRenderSemanticGroupRole.Heading2 && group.Source?.Contains("img", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(groups, group => group.Role == HtmlRenderSemanticGroupRole.Heading3 && group.Source?.Contains("span", StringComparison.OrdinalIgnoreCase) == true);
        Assert.True(groups.Count(group => group.Role == HtmlRenderSemanticGroupRole.Heading3) > 1);
        Assert.Contains(groups, group => group.Role == HtmlRenderSemanticGroupRole.Artifact && group.Source?.Contains("table", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Equal(1, tagged.StructureElements.Count(element => element.StructureType == "H3"));
        Assert.Equal(1, tagged.StructureElements.Count(element => element.StructureType == "H4"));
        Assert.Contains("/Artifact BMC", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.DoesNotContain("Decorative table", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(PdfCore.PdfInspector.Inspect(pdf).FormFields, field => field.Name == "artifact-field");
        Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields, field => field.Name == "field");
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_InlineArtifactCannotCreateABookmark() {
        const string html = "<p>Before <span style='bookmark-level:1;bookmark-label:\"Decorative\";-officeimo-pdf-tag-type:artifact'>Hidden outline</span> after</p>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Empty(rendered.Headings);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).Outlines);
        Assert.DoesNotContain(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderBookmarkAnchor>(),
            anchor => anchor.Text == "Decorative");
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_InlineBookmarkAnchorDoesNotConsumeFlowWidth() {
        const string html = "<p style='margin:0'><span id='plain'>Plain</span></p>"
            + "<p style='margin:0'><span id='bookmarked' style='bookmark-level:1'>Plain</span></p>"
            + "<p style='margin:0;width:1px'><span id='narrow-before'>X</span></p>"
            + "<p style='margin:0;width:1px'><span id='narrow-bookmark' style='bookmark-level:1'>X</span></p>"
            + "<p style='margin:0;width:1px'><span id='narrow-after'>X</span></p>";
        var options = new HtmlPdfSaveOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderText[] textVisuals = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .Where(text => text.Text == "Plain")
            .OrderBy(text => text.Y)
            .ToArray();
        Assert.Equal(2, textVisuals.Length);
        HtmlRenderText plain = textVisuals[0];
        HtmlRenderText bookmarked = textVisuals[1];

        Assert.Equal(plain.X, bookmarked.X, 6);
        HtmlRenderText[] narrow = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .Where(text => text.Text == "X")
            .OrderBy(text => text.Y)
            .ToArray();
        Assert.Equal(3, narrow.Length);
        Assert.Equal(narrow[1].Y - narrow[0].Y, narrow[2].Y - narrow[1].Y, 6);
        Assert.Equal(new[] { "Plain", "X" }, rendered.Headings.Select(heading => heading.Text).ToArray());
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_BookmarkLabelAndStateRequireAHeadingOrValidExplicitLevel() {
        const string html = "<main>"
            + "<div style='bookmark-label:\"Label only\"'>Ignored label</div>"
            + "<div style='bookmark-state:open'>Ignored state</div>"
            + "<h2 style='bookmark-label:\"Heading label\";bookmark-state:closed'>Heading text</h2>"
            + "<div style='bookmark-level:2;bookmark-label:\"Explicit entry\"'>Explicit text</div>"
            + "</main>";
        var options = new HtmlPdfSaveOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(new[] { "Heading label", "Explicit entry" }, rendered.Headings.Select(heading => heading.Text).ToArray());
        Assert.Equal(new[] { 2, 2 }, rendered.Headings.Select(heading => heading.Level).ToArray());
        Assert.Equal(2, EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderBookmarkAnchor>().Count());
        Assert.Equal(new[] { "Heading label", "Explicit entry" }, info.Outlines.Select(outline => outline.Title).ToArray());
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_BookmarkTitlesKeepLogicalSourceOrderAfterVisualRepositioning() {
        const string html = "<h1><span style='position:relative;left:100px'>A</span><span>B</span></h1>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { OutlineExpansionLevel = 64 }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal("AB", Assert.Single(rendered.Headings).Text);
        Assert.Equal("AB", Assert.Single(PdfCore.PdfInspector.Inspect(pdf).Outlines).Title);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_BookmarksKeepSourceOrderWhenDestinationsCrossPages() {
        const string html = "<style>@page{size:200px 100px;margin:0}</style>"
            + "<div style='height:120px'></div>"
            + "<h1>First source</h1>"
            + "<h1 style='position:absolute;top:0;left:0'>Second source</h1>";
        var options = new HtmlPdfSaveOptions {
            HonorCssPageRules = true,
            Margins = HtmlRenderMargins.All(0D),
            PdfOptions = new PdfCore.PdfOptions { OutlineExpansionLevel = 64 }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal(new[] { "First source", "Second source" }, rendered.Headings.Select(heading => heading.Text).ToArray());
        Assert.True(rendered.Headings[0].PageNumber > rendered.Headings[1].PageNumber);
        Assert.Equal(new[] { "First source", "Second source" }, PdfCore.PdfInspector.Inspect(pdf).Outlines.Select(outline => outline.Title).ToArray());
    }

    [Fact]
    public void HtmlPdf_DisplayContentsPreservesBookmarkAndArtifactControlsAcrossBlockChildren() {
        const string html = "<main>"
            + "<div id='bookmark-contents' style='display:contents;bookmark-level:1;bookmark-label:\"Contents entry\"'><p>Bookmark body</p><p>Second body</p></div>"
            + "<div id='artifact-contents' style='display:contents;-officeimo-pdf-tag-type:artifact'><p>Decorative first <a href='https://evotec.xyz/decorative'>link</a></p><p>Decorative second <input name='decorative-field' value='hidden'></p></div>"
            + "</main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions {
                CompressContentStreams = false,
                TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers
            }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        IReadOnlyList<HtmlRenderSemanticGroup> groups = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderSemanticGroup>().ToList();

        Assert.Contains(rendered.Headings, heading => heading.Text == "Contents entry" && heading.Level == 1);
        Assert.Contains(info.Outlines, outline => outline.Title == "Contents entry");
        Assert.Equal(2, groups.Count(group => group.Role == HtmlRenderSemanticGroupRole.Artifact && group.Source == "div#artifact-contents"));
        Assert.DoesNotContain("Decorative first", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Decorative second", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.FormFields);
        Assert.Contains("/Artifact BMC", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_DisplayContentsPreservesSemanticsForDirectFlexAndGridItems() {
        const string html = "<main>"
            + "<div style='display:flex'><div id='flex-bookmark' style='display:contents;bookmark-level:1;bookmark-label:\"Flex entry\"'><span>Flex body</span></div><div id='flex-artifact' style='display:contents;-officeimo-pdf-tag-type:artifact'><a href='https://evotec.xyz/flex'>Flex link</a><input name='flex-field' value='hidden'></div></div>"
            + "<div style='display:grid'><div id='grid-bookmark' style='display:contents;bookmark-level:1;bookmark-label:\"Grid entry\"'><span>Grid body</span></div><div id='grid-artifact' style='display:contents;-officeimo-pdf-tag-type:artifact'><a href='https://evotec.xyz/grid'>Grid link</a><input name='grid-field' value='hidden'></div></div>"
            + "</main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Contains(rendered.Headings, heading => heading.Text == "Flex entry");
        Assert.Contains(rendered.Headings, heading => heading.Text == "Grid entry");
        Assert.Contains(info.Outlines, outline => outline.Title == "Flex entry");
        Assert.Contains(info.Outlines, outline => outline.Title == "Grid entry");
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.FormFields);
        Assert.DoesNotContain("Flex link", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Grid link", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_DisplayContentsWithOnlyPositionedChildrenKeepsSemanticOwnership() {
        const string html = "<main>"
            + "<div id='block-only' style='display:contents;bookmark-level:1;-officeimo-pdf-tag-type:H2'><span style='position:fixed;left:0;top:0'>Block only</span></div>"
            + "<div style='display:flex;position:relative;height:20px'><div id='flex-only' style='display:contents;bookmark-level:1;-officeimo-pdf-tag-type:H2'><span style='position:absolute;left:0;top:0'>Flex only</span></div></div>"
            + "<div style='display:grid;position:relative;height:20px'><div id='grid-only' style='display:contents;bookmark-level:1;-officeimo-pdf-tag-type:H2'><span style='position:absolute;left:0;top:0'>Grid only</span></div></div>"
            + "</main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);

        Assert.Equal(new[] { "Block only", "Flex only", "Grid only" }, rendered.Headings.Select(heading => heading.Text).ToArray());
        Assert.Equal(new[] { "Block only", "Flex only", "Grid only" }, info.Outlines.Select(outline => outline.Title).ToArray());
        Assert.Equal(3, tagged.StructureElements.Count(element => element.StructureType == "H2"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_BookmarkFallbackLabelsUseCssNormalWhitespace() {
        const string html = "<h1>Alpha   <span>Beta</span>\n    Gamma</h1>";
        var options = new HtmlPdfSaveOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal("Alpha Beta Gamma", Assert.Single(rendered.Headings).Text);
        Assert.Equal("Alpha Beta Gamma", Assert.Single(PdfCore.PdfInspector.Inspect(pdf).Outlines).Title);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_DefaultBookmarkLabelsUseOnlyRenderedContent() {
        const string html = "<main><h1>Visible<span style='display:none'>Hidden</span><script>Script</script><span> tail</span></h1>"
            + "<h2 style='visibility:hidden'>Invisible <span style='visibility:visible'>Visible child</span></h2>"
            + "<h3>Public <span style='-officeimo-pdf-tag-type:artifact'>Decorative</span> heading</h3>"
            + "<h4 style='visibility:hidden' aria-label='Hidden metadata'>Invisible metadata</h4>"
            + "<div style='display:contents;bookmark-level:1'>Flattened<span style='display:none'>Secret</span><span> entry</span></div>"
            + "<table style='bookmark-level:1'><tr><td>Table<span style='display:none'>Private</span><span style='-officeimo-pdf-tag-type:none'>Decorative</span> entry</td></tr></table>"
            + "<input style='bookmark-level:1' aria-label='Field entry' value='Value'>"
            + "</main>";
        var options = new HtmlPdfSaveOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal(new[] { "Visible tail", "Visible child", "Public heading", "Flattened entry", "Table entry", "Field entry" }, rendered.Headings.Select(heading => heading.Text).ToArray());
        IReadOnlyList<PdfCore.PdfOutlineItem> outlines = PdfCore.PdfInspector.Inspect(pdf).Outlines;
        Assert.Equal(new[] { "Visible tail", "Flattened entry", "Table entry", "Field entry" }, outlines.Select(outline => outline.Title).ToArray());
        PdfCore.PdfOutlineItem visibleChild = Assert.Single(outlines[0].Children);
        Assert.Equal("Visible child", visibleChild.Title);
        Assert.Equal("Public heading", Assert.Single(visibleChild.Children).Title);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_DefaultBookmarkLabelsIgnoreGeneratedAndPaintTransformedText() {
        const string html = "<style>.generated::before{content:'Generated '}</style>"
            + "<h1 class='generated' style='text-transform:uppercase'>Source title</h1>"
            + "<p>Before <span class='generated' style='bookmark-level:1;text-transform:uppercase'>Inline source</span> after</p>";
        var options = new HtmlPdfSaveOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal(new[] { "Source title", "Inline source" }, rendered.Headings.Select(heading => heading.Text).ToArray());
        Assert.Equal(new[] { "Source title", "Inline source" }, PdfCore.PdfInspector.Inspect(pdf).Outlines.Select(outline => outline.Title).ToArray());
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_RepeatedFixedBookmarkAnchorsKeepOneSourceLabel() {
        const string html = "<main><h1 style='position:fixed;left:0;top:0;margin:0'>Fixed title</h1>"
            + "<p>First page</p><section style='break-before:page'>Second page</section>"
            + "<section style='break-before:page'>Third page</section></main>";
        var options = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.True(rendered.Pages.Count > 1);
        Assert.Equal("Fixed title", Assert.Single(rendered.Headings).Text);
        Assert.Equal("Fixed title", Assert.Single(PdfCore.PdfInspector.Inspect(pdf).Outlines).Title);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_OutOfFlowDescendantsRemainInsideAncestorArtifactAndBookmarkText() {
        const string html = "<main>"
            + "<div id='artifact' style='position:relative;height:60px;-officeimo-pdf-tag-type:artifact'><div style='display:contents;bookmark-level:1;-officeimo-pdf-tag-type:H2'><input style='position:absolute;left:0;top:0' name='absolute-field' value='hidden'><a style='position:fixed;left:0;top:0' href='https://evotec.xyz/fixed'>Fixed link</a></div></div>"
            + "<h1>Start<span style='position:absolute;left:120px'>Middle</span>End</h1>"
            + "</main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);

        Assert.Equal("StartMiddleEnd", Assert.Single(rendered.Headings).Text);
        Assert.Equal("StartMiddleEnd", Assert.Single(info.Outlines).Title);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.FormFields);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "H2");
        Assert.DoesNotContain("Fixed link", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_OutOfFlowArtifactRemainsInsideOuterFlattenedBookmark() {
        const string html = "<main>"
            + "<div style='display:contents;bookmark-level:1;bookmark-label:\"Outer bookmark\";-officeimo-pdf-tag-type:H2'>"
            + "<div style='display:contents;-officeimo-pdf-tag-type:artifact'>"
            + "<a style='display:block;position:fixed;left:0;top:0' href='https://evotec.xyz/hidden'>Hidden artifact link</a>"
            + "</div></div></main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);

        Assert.Equal("Outer bookmark", Assert.Single(rendered.Headings).Text);
        Assert.Equal("Outer bookmark", Assert.Single(info.Outlines).Title);
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "H2");
        Assert.Empty(info.LinkAnnotations);
        Assert.DoesNotContain("Hidden artifact link", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("fixed")]
    [InlineData("absolute")]
    public void HtmlPdf_NestedArtifactsSuppressIntermediateFlattenedBookmark(string position) {
        string html = "<main style='position:relative'>"
            + "<div style='display:contents;-officeimo-pdf-tag-type:artifact'>"
            + "<div style='display:contents;bookmark-level:1;bookmark-label:\"Hidden bookmark\";-officeimo-pdf-tag-type:H2'>"
            + "<div style='display:contents;-officeimo-pdf-tag-type:artifact'>"
            + "<a style='display:block;position:" + position + ";left:0;top:0' href='https://evotec.xyz/hidden'>Hidden nested artifact link</a>"
            + "</div></div></div></main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);

        Assert.Empty(rendered.Headings);
        Assert.Empty(info.Outlines);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "H2");
        Assert.Empty(info.LinkAnnotations);
        Assert.DoesNotContain("Hidden nested artifact link", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_PagedDisplayContentsRetainsListStructureAndBookmarkAnchor() {
        string longText = string.Join(" ", Enumerable.Repeat("breakable", 80));
        string html = "<main><ol style='display:contents;bookmark-level:1;bookmark-label:\"Paged list\"'><li style='display:contents'><p style='width:120px;font-size:16px;line-height:20px;margin:0'>" + longText + "</p><p style='margin:0'>Continuation body</p></li></ol></main>";
        var options = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 1.25D),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);

        Assert.True(rendered.Pages.Count > 1);
        Assert.Equal("Paged list", Assert.Single(rendered.Headings).Text);
        PdfCore.PdfStructureElementInfo list = Assert.Single(tagged.StructureElements, element => element.StructureType == "L");
        PdfCore.PdfStructureElementInfo listItem = Assert.Single(tagged.StructureElements, element => element.StructureType == "LI");
        PdfCore.PdfStructureElementInfo listBody = Assert.Single(tagged.StructureElements, element => element.StructureType == "LBody");
        Assert.Contains(listItem.ObjectNumber, list.ChildElementObjectNumbers);
        Assert.Contains(listBody.ObjectNumber, listItem.ChildElementObjectNumbers);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_DisplayContentsReusesOneListStructureAcrossSiblingFragments() {
        const string html = "<main><ol style='display:contents'><li>First item body</li><li>Second item body</li></ol><ol><li style='display:contents'><p>Flattened item body</p><p>Continuation body</p></li></ol></main>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);

        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "L"));
        Assert.Equal(3, tagged.StructureElements.Count(element => element.StructureType == "LI"));
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "Lbl"));
        Assert.Equal(3, tagged.StructureElements.Count(element => element.StructureType == "LBody"));
    }

    [Fact]
    public void HtmlRender_InvalidPdfControlsOnSpecializedAndInlineElementsAreNeverSilent() {
        const string html = "<hr style='-officeimo-pdf-tag-type:made-up;bookmark-state:sideways'><p><span style='-officeimo-pdf-tag-type:also-made-up;bookmark-state:sideways'>Text</span></p><div style='bookmark-level:bogus'>Not a heading</div>";

        HtmlRenderDocument permissive = HtmlRenderTestDriver.Render(html);
        Assert.Equal(2, permissive.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PdfSemanticTagUnsupported));
        Assert.Equal(3, permissive.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BookmarkValueUnsupported));
        Assert.DoesNotContain(permissive.Headings, heading => heading.Text.Contains("Not a heading", StringComparison.Ordinal));

        HtmlConversionException exception = Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        }));
        Assert.Contains(exception.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PdfSemanticTagUnsupported);
        Assert.Contains(exception.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BookmarkValueUnsupported);
    }

    [Fact]
    public void HtmlPdf_PagedGeneratedMarginTextIsDecorativeArtifact() {
        const string html = "<style>@page{size:320px 180px;margin:32px;@top-left{content:\"Page \" counter(page)}}</style><p>Body</p>";
        var options = new HtmlPdfSaveOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            PdfOptions = new PdfCore.PdfOptions { CompressContentStreams = false }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        HtmlRenderSemanticGroup artifact = Assert.Single(rendered.Pages[0].Scene.OfType<HtmlRenderSemanticGroup>(), group =>
            group.Role == HtmlRenderSemanticGroupRole.Artifact && group.Source?.Contains("top-left", StringComparison.Ordinal) == true);

        Assert.Contains(artifact.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Page 1");
        Assert.Contains("/Artifact BMC", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.DoesNotContain("Page 1", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Page 1", PdfCore.PdfReadDocument.Open(pdf, new PdfCore.PdfReadOptions { IncludeArtifactText = true }).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_UsesXmlLanguageWhenHtmlLanguageIsEmpty() {
        const string html = "<html lang='' xml:lang='fr-FR'><body><p>Langue</p></body></html>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());

        Assert.Equal("fr-FR", rendered.Metadata.Language);
        Assert.Equal("fr-FR", PdfCore.PdfReadDocument.Open(pdf).CatalogLanguage);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_PreservesCallerDocumentLanguageOverHtmlMetadata() {
        const string html = "<html lang='fr-FR'><body><p>Language precedence</p></body></html>";
        var options = new HtmlPdfSaveOptions();
        options.PdfOptions.Language = "en-US";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Equal("en-US", PdfCore.PdfReadDocument.Open(pdf).CatalogLanguage);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_UsesSharedPagedLayoutAndPreservesTextAndLink() {
        const string linkUri = "https://example.test/direct-pdf";
        string html = """
            <style>@media print { h1 { color:#224466; } }</style>
            <h1>RenderedPdfMarker</h1>
            <p><a href="https://example.test/direct-pdf">RenderedLinkMarker</a></p>
            <div style="break-before:page"><p>SecondPageMarker</p></div>
            """;
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();
        options.PageSize = new OfficePageSize(4D, 3D);
        options.Margins = HtmlRenderMargins.All(20D);

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(2, info.PageCount);
        Assert.Contains("RenderedPdfMarker", text, StringComparison.Ordinal);
        Assert.Contains("RenderedLinkMarker", text, StringComparison.Ordinal);
        Assert.Contains("SecondPageMarker", text, StringComparison.Ordinal);
        Assert.Contains(linkUri, info.LinkUris);
        Assert.Equal(HtmlRenderMode.Paged, options.Mode);
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_UsesManagedFontFallbacksForUnicodeText() {
        const string marker = "Café Ω Ж שלום سلام";
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse("<p>" + marker + "</p>").ToPdf(options);
        string extracted = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(PdfCore.PdfTextFallbackFeatures.Default, options.TextFallbacks);
        Assert.Equal(PdfCore.PdfTextShapingMode.LatinLigatures, options.TextShapingMode);
        Assert.Contains(marker, extracted, StringComparison.Ordinal);
        var fallbackProbe = new PdfCore.PdfOptions();
        if (fallbackProbe.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            Assert.True(PdfCore.PdfDiagnostics.Analyze(pdf).EmbeddedFontCount > 0);
        }
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_UsesRegularFallbackCoverageWhenBoldSystemFaceIsNarrower() {
        const string marker = "Bold שלום سلام";

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse("<h1>" + marker + "</h1>").ToPdf(new HtmlPdfSaveOptions());

        Assert.Contains(marker, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_PreservesCallerUnicodeFontWhenManagedFallbacksAreActive() {
        if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfCore.PdfEmbeddedFontFamily? installed) || installed == null) return;
        const string marker = "Caller שלום سلام";
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();
        options.FontFamily = new PdfCore.PdfEmbeddedFontFamily("CallerUnicode", installed.Regular);

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse("<p>" + marker + "</p>").ToPdf(options);
        PdfCore.PdfDiagnosticReport report = PdfCore.PdfDiagnostics.Analyze(pdf);

        Assert.Contains(marker, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(report.Fonts, font =>
            font.HasEmbeddedFontFile
            && font.BaseFont?.Contains("CallerUnicode", StringComparison.OrdinalIgnoreCase) == true);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_LoadsManagedFontFallbacksOnlyWhenSceneTextRequiresUnicode() {
        HtmlRenderDocument winAnsi = HtmlRenderTestDriver.Render("<p>Invoice Café — paid</p>");
        HtmlRenderDocument unicode = HtmlRenderTestDriver.Render("<p>Invoice Ω Ж שלום سلام</p>");

        Assert.Equal(
            PdfCore.PdfTextFallbackFeatures.None,
            HtmlPdfRenderedConverter.ResolveTextFallbackFeatures(winAnsi, PdfCore.PdfTextFallbackFeatures.Default));
        Assert.Equal(
            PdfCore.PdfTextFallbackFeatures.Default,
            HtmlPdfRenderedConverter.ResolveTextFallbackFeatures(unicode, PdfCore.PdfTextFallbackFeatures.Default));
        Assert.Equal(
            PdfCore.PdfTextFallbackFeatures.None,
            HtmlPdfRenderedConverter.ResolveTextFallbackFeatures(unicode, PdfCore.PdfTextFallbackFeatures.None));
    }

    [Fact]
    public void HtmlRenderer_PositionsSimpleRtlTextAndDiagnosesOnlyRemainingBidiStages() {
        const string html = "<div style='width:200px'><p id='declared' dir='rtl'>Latin text</p><p id='hebrew' dir='rtl'>שלום 123</p><h2 id='arabic' dir='rtl'>سلام</h2><p id='authored' dir='rtl'>\uFE8F\uFE8F</p><p id='syriac' dir='rtl'>ܫܠܡ</p><p id='control'>abc\u202Edef</p></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        IReadOnlyList<HtmlRenderText> text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();
        IReadOnlyList<HtmlRenderText> hebrew = text
            .Where(run => run.Text.Length == 1 && "שלום".Contains(run.Text, StringComparison.Ordinal))
            .OrderBy(run => run.PaintOrder)
            .ToList();

        Assert.Equal(4, hebrew.Count);
        Assert.Equal("שלום", string.Concat(hebrew.Select(run => run.Text)));
        HtmlRenderLogicalTextGroup logicalGroup = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            group => group.Text == "שלום 123");
        Assert.Equal("שלום 123", logicalGroup.Text);
        for (int index = 1; index < hebrew.Count; index++) Assert.True(hebrew[index].X < hebrew[index - 1].X);
        HtmlRenderText number = Assert.Single(text, run => run.Text == "123");
        Assert.Equal("שלום 123", string.Concat(text.Where(run => Math.Abs(run.Y - number.Y) < 0.001D).OrderBy(run => run.PaintOrder).Select(run => run.Text)));
        Assert.True(number.X < hebrew.Min(run => run.X));

        HtmlRenderLogicalTextGroup arabicGroup = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            group => group.Text == "سلام");
        Assert.Equal("\uFEB3\uFEE0\uFE8E\uFEE1", string.Concat(arabicGroup.Visuals.OfType<HtmlRenderText>().Select(run => run.Text)));
        HtmlRenderLogicalTextGroup authoredForms = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            group => group.Text == "\uFE8F\uFE8F");
        Assert.Equal("\uFE91\uFE90", string.Concat(authoredForms.Visuals.OfType<HtmlRenderText>().Select(run => run.Text)));
        Assert.Contains("سلام", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("\uFEB3\uFEE0\uFE8E\uFEE1", rendered.Text, StringComparison.Ordinal);
        HtmlRenderHeading arabicHeading = Assert.Single(rendered.Headings, heading => heading.Level == 2);
        Assert.Equal("سلام", arabicHeading.Text);
        Assert.True(HtmlConversionDocument.Parse(html).ToPng().Length > 8);
        string svg = HtmlConversionDocument.Parse(html).ToSvg();
        Assert.All("\uFEB3\uFEE0\uFE8E\uFEE1", character => Assert.Contains(character.ToString(), svg, StringComparison.Ordinal));

        HtmlRenderLogicalTextGroup controlGroup = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            group => group.Text == "abcdef");
        Assert.Equal("abcdef", controlGroup.Text);
        IReadOnlyList<HtmlRenderText> overridden = controlGroup.Visuals.OfType<HtmlRenderText>().Where(run => "def".Contains(run.Text, StringComparison.Ordinal)).ToList();
        Assert.Equal(3, overridden.Count);
        Assert.True(overridden[1].X < overridden[0].X);
        Assert.True(overridden[2].X < overridden[1].X);
        HtmlDiagnostic bidiDiagnostic = Assert.Single(
            rendered.Diagnostics,
            diagnostic =>
                diagnostic.Code == HtmlRenderDiagnosticCodes.BidiLayoutUnsupported ||
                diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Equal(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, bidiDiagnostic.Code);
        Assert.Equal("p#syriac", bidiDiagnostic.Source);
        Assert.DoesNotContain(HtmlRenderDiagnosticCodes.BidiLayoutUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.False(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BidiLayoutUnsupported, out _));
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, out _));
    }

    [Fact]
    public void HtmlRenderer_PositionsHebrewRunInsideLtrTextWithoutChangingLogicalSceneOrder() {
        const string html = "<p style='margin:0;width:240px'>Left שלום 42</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        IReadOnlyList<HtmlRenderText> runs = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().OrderBy(run => run.PaintOrder).ToList();
        IReadOnlyList<HtmlRenderText> hebrew = runs.Where(run => run.Text.Length == 1 && "שלום".Contains(run.Text, StringComparison.Ordinal)).ToList();
        HtmlRenderText left = Assert.Single(runs, run => run.Text.Contains("Left", StringComparison.Ordinal));
        HtmlRenderText number = Assert.Single(runs, run => run.Text.Contains("42", StringComparison.Ordinal));

        Assert.Equal("Left שלום 42", string.Concat(runs.Select(run => run.Text)));
        Assert.Equal(4, hebrew.Count);
        Assert.True(left.X < hebrew.Min(run => run.X));
        Assert.True(number.X > hebrew.Max(run => run.X));
        for (int index = 1; index < hebrew.Count; index++) Assert.True(hebrew[index].X < hebrew[index - 1].X);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BidiLayoutUnsupported || diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.True(HtmlConversionDocument.Parse(html).ToPng().Length > 8);
        Assert.Contains("ש", HtmlConversionDocument.Parse(html).ToSvg(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRenderer_MirrorsPairedPunctuationInSimpleRtlRuns() {
        const string html = "<p dir='rtl' style='margin:0;width:240px'>(אבג)</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderLogicalTextGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            item => item.Text == "(אבג)");
        string visual = string.Concat(group.Visuals
            .OfType<HtmlRenderText>()
            .OrderBy(static item => item.X)
            .Select(static item => item.Text));

        Assert.Equal("(גבא)", visual);
        Assert.Equal("(אבג)", group.Text);
    }

    [Fact]
    public void HtmlRenderer_PositionsNestedLtrIsolateRunBeforeItsHebrewRun() {
        const string html = "<p style='margin:0;width:240px'>A\u2067שלום abc\u2069B</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderLogicalTextGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            item => item.Text == "Aשלום abcB");
        HtmlRenderText latin = Assert.Single(group.Visuals.OfType<HtmlRenderText>(), run => run.Text.Contains("abc", StringComparison.Ordinal));
        IReadOnlyList<HtmlRenderText> hebrew = group.Visuals
            .OfType<HtmlRenderText>()
            .Where(run => run.Text.Length == 1 && "שלום".Contains(run.Text, StringComparison.Ordinal))
            .ToList();

        Assert.Equal("Aשלום abcB", group.Text);
        Assert.Equal(4, hebrew.Count);
        Assert.True(latin.X < hebrew.Min(static run => run.X));
    }

    [Fact]
    public void HtmlRenderer_CarriesBidiOverridesAcrossInlineFormattingBoundaries() {
        const string html = "<p style='margin:0'><span>\u202E</span><b>abc</b><span>\u202C</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderLogicalTextGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            item => item.Text == "abc" && item.Visuals.OfType<HtmlRenderText>().Any());
        HtmlRenderText text = Assert.Single(group.Visuals.OfType<HtmlRenderText>());

        Assert.Equal("cba", text.Text);
        Assert.Equal("abc", group.Text);
    }

    [Fact]
    public void HtmlRenderer_DoesNotReorderResolvedBidiTextAgainInDrawingExport() {
        const string html = "<p style='margin:0'><span>\u202B</span><b>אבג</b><span>\u202C</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderLogicalTextGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            item => item.Text == "אבג" && item.Visuals.OfType<HtmlRenderText>().Any());
        HtmlRenderText text = Assert.Single(group.Visuals.OfType<HtmlRenderText>());
        string svg = OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());

        Assert.Equal("גבא", text.Text);
        Assert.Contains("\u202Dגבא\u202C", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("\u202Dאבג\u202C", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRenderer_ReportsStackedInlineTextOnceInLogicalOutput() {
        const string html = "<p style='margin:0'><span>\u202B</span>Start <span style='position:relative;z-index:1'>אבג</span> End<span>\u202C</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));

        Assert.Equal(1, rendered.Text.Split(new[] { "אבג" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void HtmlRenderer_RetainsLogicalTextOrderAcrossVisuallyReorderedStyledRuns() {
        const string html = "<p style='margin:0'><span>\u202E</span><b>abc</b><i>def</i><span>\u202C</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        IReadOnlyList<HtmlRenderLogicalTextGroup> groups = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderLogicalTextGroup>()
            .Where(group => group.Text is "abc" or "def")
            .ToList();

        Assert.Equal(2, groups.Count);
        Assert.Contains("abcdef", rendered.Text, StringComparison.Ordinal);
        HtmlRenderText abc = Assert.Single(Assert.Single(groups, static group => group.Text == "abc").Visuals.OfType<HtmlRenderText>());
        HtmlRenderText def = Assert.Single(Assert.Single(groups, static group => group.Text == "def").Visuals.OfType<HtmlRenderText>());
        Assert.True(def.X < abc.X);
        Assert.Equal("cba", abc.Text);
        Assert.Equal("fed", def.Text);
    }

    [Fact]
    public void HtmlRenderer_CarriesBidiOverridesAcrossWrappedLines() {
        const string html = "<p style='margin:0;width:32px;font-size:12px'>\u202Eone two four\u202C</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        IReadOnlyList<HtmlRenderLogicalTextGroup> groups = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderLogicalTextGroup>()
            .ToList();

        Assert.Contains(groups, group =>
            group.Text == "one" && group.Visuals.OfType<HtmlRenderText>().SingleOrDefault()?.Text == "eno");
        Assert.Contains(groups, group =>
            group.Text == "two" && group.Visuals.OfType<HtmlRenderText>().SingleOrDefault()?.Text == "owt");
        Assert.Contains(groups, group =>
            group.Text == "four" && group.Visuals.OfType<HtmlRenderText>().SingleOrDefault()?.Text == "ruof");
    }

    [Fact]
    public void HtmlRenderer_ResolvesLogicalTextAlignmentAgainstElementDirection() {
        const string html = "<div style='width:160px'><p id='start' dir='rtl' style='margin:0'>Start</p><p id='end' dir='rtl' style='margin:0;text-align:end'>End</p><p id='left' dir='rtl' style='margin:0;text-align:left'>Left</p></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();

        HtmlRenderText start = Assert.Single(text, item => item.Text == "Start");
        HtmlRenderText end = Assert.Single(text, item => item.Text == "End");
        HtmlRenderText left = Assert.Single(text, item => item.Text == "Left");
        Assert.True(start.X > 100D);
        Assert.Equal(0D, end.X, 6);
        Assert.Equal(0D, left.X, 6);
    }

    [Fact]
    public void HtmlRenderer_MatchParentUsesPhysicalParentAlignment() {
        const string html =
            "<div style='width:200px;text-align:center'><p style='margin:0;text-align:match-parent'>Matched</p></div>";

        var parsed = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(parsed);
        Assert.Equal("center", styles[parsed.QuerySelector("div")!].GetValue("text-align"));
        Assert.Equal("match-parent", styles[parsed.QuerySelector("p")!].GetValue("text-align"));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions {
                ViewportWidth = 200D,
                Margins = HtmlRenderMargins.All(0D)
            });

        HtmlRenderText text = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text == "Matched");
        // Inline text fragments are already physically positioned, so their
        // local drawing frame remains left-aligned. The inherited alignment is
        // observable through the centered physical position.
        Assert.True(text.X > 50D);
    }

    [Fact]
    public void HtmlRenderer_TableCellMatchParentUsesRowAlignment() {
        const string html =
            "<table style='width:200px;text-align:left'><tr style='text-align:center'><td style='text-align:match-parent'>MatchedCell</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions {
                ViewportWidth = 200D,
                Margins = HtmlRenderMargins.All(0D)
            });

        HtmlRenderText text = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text == "MatchedCell");
        Assert.True(text.X > 40D);
    }

    [Fact]
    public void HtmlRenderer_RootDirectionMetadataHonorsCssInitialReset() {
        const string html =
            "<!doctype html><html dir='rtl' style='direction:initial'><body><p>LTR metadata</p></body></html>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html));

        Assert.Equal(HtmlRenderTextDirection.LeftToRight, rendered.Metadata.Direction);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_TagsRasterAndVectorImageAlternativeTextAsFigures() {
        string rasterData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        const string vectorData = "%3Csvg xmlns='http://www.w3.org/2000/svg' width='2' height='2'%3E%3Crect width='2' height='2' fill='red'/%3E%3C/svg%3E";
        string html = "<img alt='Raster badge' width='24' height='24' src='data:image/png;base64," + rasterData + "'>"
            + "<img alt='Vector badge' width='24' height='24' src=\"data:image/svg+xml," + vectorData + "\">";

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        IReadOnlyList<PdfCore.PdfStructureElementInfo> figures = tagged.StructureElements
            .Where(element => element.StructureType == "Figure")
            .ToList();

        Assert.Equal(2, figures.Count);
        Assert.Contains(figures, figure => figure.AlternateText == "Raster badge");
        Assert.Contains(figures, figure => figure.AlternateText == "Vector badge");
        Assert.True(tagged.FiguresHaveAlternateText);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_PreservesListItemLabelAndBodySemantics() {
        const string html = "<ol><li>First item</li><li>Second item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderSemanticGroup listScene = Assert.Single(rendered.Pages[0].Scene.OfType<HtmlRenderSemanticGroup>());
        Assert.Equal(HtmlRenderSemanticGroupRole.List, listScene.Role);
        IReadOnlyList<HtmlRenderSemanticGroup> items = listScene.Visuals
            .OfType<HtmlRenderSemanticGroup>()
            .Where(group => group.Role == HtmlRenderSemanticGroupRole.ListItem)
            .ToList();
        Assert.Equal(2, items.Count);
        Assert.All(items, item => {
            Assert.Contains(item.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.ListLabel);
            Assert.Contains(item.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.ListBody);
        });

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        PdfCore.PdfStructureElementInfo list = Assert.Single(tagged.StructureElements, element => element.StructureType == "L");
        IReadOnlyList<PdfCore.PdfStructureElementInfo> pdfItems = tagged.StructureElements.Where(element => element.StructureType == "LI").ToList();
        Assert.Equal(2, pdfItems.Count);
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "Lbl"));
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "LBody"));
        Assert.All(pdfItems, item => Assert.Contains(item.ObjectNumber, list.ChildElementObjectNumbers));
        Assert.Contains("1. First item", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("2. Second item", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_PreservesNestedTableCaptionRowAndCellSemantics() {
        const string html = "<table><caption>Quarterly status</caption><tr><th scope='row' rowspan='2'>Area</th><th colspan='2'>Status</th></tr><tr><td>Green</td><td>Ready</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderSemanticGroup tableScene = Assert.Single(rendered.Pages[0].Scene.OfType<HtmlRenderSemanticGroup>());
        Assert.Equal(HtmlRenderSemanticGroupRole.Table, tableScene.Role);
        Assert.Contains(tableScene.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Caption);
        Assert.Equal(2, tableScene.Visuals.OfType<HtmlRenderSemanticGroup>().Count(group => group.Role == HtmlRenderSemanticGroupRole.TableRow));

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);
        PdfCore.PdfStructureElementInfo table = Assert.Single(tagged.StructureElements, element => element.StructureType == "Table");
        PdfCore.PdfStructureElementInfo caption = Assert.Single(tagged.StructureElements, element => element.StructureType == "Caption");
        IReadOnlyList<PdfCore.PdfStructureElementInfo> rows = tagged.StructureElements.Where(element => element.StructureType == "TR").ToList();
        Assert.Equal(2, rows.Count);
        Assert.Contains(caption.ObjectNumber, table.ChildElementObjectNumbers);
        Assert.All(rows, row => Assert.Contains(row.ObjectNumber, table.ChildElementObjectNumbers));
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "TH"));
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "TD"));
        string raw = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/Scope /Row", raw, StringComparison.Ordinal);
        Assert.Contains("/ColSpan 2", raw, StringComparison.Ordinal);
        Assert.Contains("/RowSpan 2", raw, StringComparison.Ordinal);
        Assert.Contains("Quarterly status", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRenderDiagnostics_AreAllRegisteredInThePublicCatalog() {
        Assert.Contains(HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.All(HtmlRenderDiagnosticCodes.All, code =>
            Assert.True(HtmlDiagnosticCatalog.TryGet(code, out _), code));
    }

    private static int CountRenderedTextLines(HtmlRenderPage page) =>
        page.Visuals.OfType<HtmlRenderText>()
            .Where(text => text.Text.StartsWith("word", StringComparison.Ordinal))
            .Select(text => Math.Round(text.Y, 3))
            .Distinct()
            .Count();
}
