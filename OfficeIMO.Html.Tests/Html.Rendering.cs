using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
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
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(1D));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            HtmlRenderTestDriver.RenderAsync(HtmlConversionDocument.Parse(html), new HtmlRenderOptions { ViewportWidth = 240D }, cancellation.Token));
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

        Assert.Contains("AsyncPdfMarker", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
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
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Load(pdf);
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
        options.MaxResourceBytes = 1024L;
        options.MaxTotalResourceBytes = 4096L;
        options.MaxResourceCount = 12;
        options.MaxStylesheetImportDepth = 4;
        options.UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        options.ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(null);

        HtmlPdfResourcePolicySummary summary = options.GetResourcePolicySummary();

        Assert.True(summary.HasResourceResolver);
        Assert.True(summary.AllowSystemFontEmbedding);
        Assert.False(summary.AllowLocalFileAccess);
        Assert.False(summary.AllowRemoteResourceResolution);
        Assert.True(summary.AllowDataUris);
        Assert.True(summary.AllowEmbeddedPackageResources);
        Assert.Equal(TimeSpan.FromSeconds(5D), summary.ResourceTimeout);
        Assert.Equal(1024L, summary.MaxResourceBytes);
        Assert.Equal(4096L, summary.MaxTotalResourceBytes);
        Assert.Equal(12, summary.MaxResourceCount);
        Assert.Equal(4, summary.MaxStylesheetImportDepth);
        Assert.Contains("https", summary.AllowedUrlSchemes);
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

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText().Replace("\r", string.Empty).Replace("\n", string.Empty);
        Assert.Contains("MhtmlPdfMarker", text, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
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
    public void MhtmlPdf_ExposesCompleteDirectLifecycle() {
        MethodInfo[] methods = typeof(HtmlPdfConverterExtensions)
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
        string pdfText = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
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
        string pdfText = PdfCore.PdfReadDocument.Load(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
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
        string pdfText = PdfCore.PdfReadDocument.Load(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
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
        string pdfText = PdfCore.PdfReadDocument.Load(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
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
        string pdfText = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        Assert.Equal(rendered.Pages.Count, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Contains("FirstPage", pdfText, StringComparison.Ordinal);
        Assert.Contains("Page 2 of " + rendered.Pages.Count, pdfText, StringComparison.Ordinal);
        Assert.Contains("L2", pdfText, StringComparison.Ordinal);
        Assert.Contains("R3", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_Paged_DiagnosesPseudoPageGeometryUntilPerPageReflowIsAvailable() {
        string html = "<style>@page { size:3in 2in; margin:0.25in; } @page :first { size:2in 2in; margin:0.5in; @top-left { content:\"First\"; } }</style><p>Body</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(288D, page.Width, 3);
        Assert.Equal(192D, page.Height, 3);
        Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "First");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PageSelectorPending);
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
        string pdfText = PdfCore.PdfReadDocument.Load(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
        Assert.Contains("Invoice", pdfText, StringComparison.Ordinal);
        Assert.Contains("IL", pdfText, StringComparison.Ordinal);
        Assert.Contains("Report", pdfText, StringComparison.Ordinal);
        Assert.Contains("ReportBody", pdfText, StringComparison.Ordinal);
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
        string pdfText = PdfCore.PdfReadDocument.Load(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText();
        foreach (string marker in markers) Assert.Contains(marker, pdfText, StringComparison.Ordinal);
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
        string pdfText = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
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
        Assert.Contains("StableMarker", PdfCore.PdfReadDocument.Load(firstPdf).ExtractText(), StringComparison.Ordinal);
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
        const string html = "<!doctype html><html lang='pl-PL' dir='rtl'><head><title>Semantic document</title></head><body><main><h1>Semantic <em>heading</em></h1><p>Semantic <strong>paragraph</strong>.</p><h2>Nested detail</h2></main></body></html>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);
        Assert.Equal("Semantic document", rendered.Metadata.Title);
        Assert.Equal("pl-PL", rendered.Metadata.Language);
        Assert.Equal(HtmlRenderTextDirection.RightToLeft, rendered.Metadata.Direction);
        Assert.Equal("Semantic document", info.Metadata.Title);
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
        Assert.Contains("Semantic heading", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
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
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

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
        string extracted = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

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

        Assert.Contains(marker, PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_DirectRenderer_PreservesCallerUnicodeFontWhenManagedFallbacksAreActive() {
        if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfCore.PdfEmbeddedFontFamily? installed) || installed == null) return;
        const string marker = "Caller שלום سلام";
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions();
        options.FontFamily = new PdfCore.PdfEmbeddedFontFamily("CallerUnicode", installed.Regular);

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse("<h1>" + marker + "</h1>").ToPdf(options);
        PdfCore.PdfDiagnosticReport report = PdfCore.PdfDiagnostics.Analyze(pdf);

        Assert.Contains(marker, PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(report.Fonts, font => font.BaseFont?.Contains("CallerUnicode", StringComparison.Ordinal) == true);
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

        Assert.Collection(
            rendered.Diagnostics
                .Where(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BidiLayoutUnsupported || diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported)
                .OrderBy(diagnostic => diagnostic.Source),
            diagnostic => {
                Assert.Equal(HtmlRenderDiagnosticCodes.BidiLayoutUnsupported, diagnostic.Code);
                Assert.Equal("p#control", diagnostic.Source);
            },
            diagnostic => {
                Assert.Equal(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, diagnostic.Code);
                Assert.Equal("p#syriac", diagnostic.Source);
            });
        Assert.Contains(HtmlRenderDiagnosticCodes.BidiLayoutUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BidiLayoutUnsupported, out _));
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, out _));
    }

    [Fact]
    public void HtmlRenderer_PositionsHebrewRunInsideLtrTextWithoutChangingLogicalSceneOrder() {
        const string html = "<p style='margin:0;width:240px'>Left שלום 42</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        IReadOnlyList<HtmlRenderText> runs = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().OrderBy(run => run.PaintOrder).ToList();
        IReadOnlyList<HtmlRenderText> hebrew = runs.Where(run => run.Text.Length == 1 && "שלום".Contains(run.Text, StringComparison.Ordinal)).ToList();
        HtmlRenderText left = Assert.Single(runs, run => run.Text == "Left ");
        HtmlRenderText number = Assert.Single(runs, run => run.Text == "42");

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
        Assert.Contains("1. First item", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("2. Second item", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
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
        Assert.Contains("Quarterly status", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRenderDiagnostics_AreAllRegisteredInThePublicCatalog() {
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
