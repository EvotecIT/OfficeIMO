using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public async Task HtmlResourceSession_OwnsDedupMimeBudgetsCacheAndDigestEvidence() {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<img src='https://assets.example.test/chart.png'><img src='https://assets.example.test/chart.png'>");
        int requests = 0;
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => {
                requests++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(png, "image/png"));
            },
            MaxResourceBytes = 1024,
            MaxTotalResourceBytes = 2048,
            MaxResourceCount = 2,
            MaxResourceRequests = 2
        };

        HtmlResourceSession session = await HtmlResourceSession.ResolveAsync(source.ResourceManifest, options);

        Assert.Equal(1, requests);
        Assert.Equal(1, session.ResolverRequestCount);
        Assert.Equal(1, session.AcceptedResourceCount);
        Assert.Equal(png.Length, session.AcceptedResourceBytes);
        HtmlResourceSessionEntry entry = Assert.Single(session.Resources);
        Assert.Equal("https://assets.example.test/chart.png", entry.CanonicalSource);
        Assert.Equal("image/png", entry.ContentType);
        Assert.Equal(64, entry.Sha256.Length);
        Assert.True(session.TryGet("https://assets.example.test/chart.png", null, out HtmlResolvedResource cached));
        Assert.Equal(png, cached.Bytes);

        options.ResourceResolver = (request, cancellationToken) =>
            Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(new byte[] { 1, 2, 3 }, "text/plain"));
        HtmlResourceSession rejected = await HtmlResourceSession.ResolveAsync(source.ResourceManifest, options);
        Assert.Empty(rejected.Resources);
        Assert.Contains(rejected.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceContentTypeRejected);
    }

    [Fact]
    public async Task HtmlRender_ResolvesResourcesConcurrentlyWithinTheConfiguredBound() {
        int active = 0;
        int maximumActive = 0;
        int calls = 0;
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var options = new HtmlRenderOptions {
            MaxConcurrentResourceLoads = 2,
            ResourceResolver = async (request, cancellationToken) => {
                Interlocked.Increment(ref calls);
                int current = Interlocked.Increment(ref active);
                int observed;
                do {
                    observed = maximumActive;
                    if (observed >= current) break;
                } while (Interlocked.CompareExchange(ref maximumActive, current, observed) != observed);
                await Task.Delay(30, cancellationToken);
                Interlocked.Decrement(ref active);
                return new HtmlResolvedResource(png, "image/png");
            },
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        await HtmlRenderTestDriver.RenderAsync(
            string.Concat(Enumerable.Range(1, 4).Select(index => "<img src='https://assets.example.test/" + index + ".png' width='2' height='2'>")),
            options);

        Assert.Equal(4, calls);
        Assert.InRange(maximumActive, 2, options.MaxConcurrentResourceLoads);
    }

    [Fact]
    public async Task HtmlRender_BoundsResolverMissesWithTheRequestBudget() {
        int calls = 0;
        var options = new HtmlRenderOptions {
            MaxConcurrentResourceLoads = 2,
            MaxResourceCount = 1,
            MaxResourceRequests = 2,
            ResourceResolver = (request, cancellationToken) => {
                Interlocked.Increment(ref calls);
                return Task.FromResult<HtmlResolvedResource?>(null);
            },
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<img src='https://assets.example.test/1.png'><img src='https://assets.example.test/2.png'><img src='https://assets.example.test/3.png'>",
            options);

        Assert.Equal(2, calls);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded);
    }

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
    public async Task HtmlRender_IntersectsRenderPoliciesWithUntrustedDocumentPolicies() {
        const string source = "file:///approved/image.png";
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<a href='file:///secret/report.pdf'>Local report</a><img src='" + source + "' width='2' height='2'>");
        var resourcePolicy = HtmlUrlPolicy.CreateOfficeIMOProfile();
        resourcePolicy.DisallowFileUrls = false;
        int resolverCalls = 0;
        var options = new HtmlRenderOptions {
            UrlPolicy = resourcePolicy,
            ResourceUrlPolicy = resourcePolicy,
            ResourceResolver = (request, cancellationToken) => {
                resolverCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(png, "image/png"));
            },
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Visuals).ToList();
        Assert.Equal(0, resolverCalls);
        Assert.Empty(visuals.OfType<HtmlRenderImage>());
        Assert.DoesNotContain(visuals.OfType<HtmlRenderText>(), text =>
            text.LinkUri?.StartsWith("file:", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == "HyperlinkRejectedByPolicy");
    }

    [Fact]
    public async Task HtmlRender_TrustedDocumentAllowsExplicitFileResourcePolicy() {
        const string source = "file:///approved/image.png";
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<img src='" + source + "' width='2' height='2'>",
            HtmlConversionDocumentOptions.CreateTrustedProfile());
        var resourcePolicy = HtmlUrlPolicy.CreateOfficeIMOProfile();
        resourcePolicy.DisallowFileUrls = false;
        var options = new HtmlRenderOptions {
            ResourceUrlPolicy = resourcePolicy,
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                new HtmlResolvedResource(png, "image/png")),
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Visuals), visual => visual is HtmlRenderImage);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy");
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
    public void HtmlRender_InvalidInlineFontMimeDoesNotConsumeImageBudgetOrMislabelTheResource() {
        byte[] font = CreateHtmlRenderTestFont();
        string png = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<style>@font-face{font-family:Rejected;src:url('data:image/png;base64,"
            + Convert.ToBase64String(font)
            + "')}p{font-family:Rejected}</style><p>Fallback</p><img src='data:image/png;base64,"
            + png + "'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            MaxResourceCount = 1,
            MaxResourceBytes = 1024L * 1024L,
            MaxTotalResourceBytes = 2L * 1024L * 1024L,
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Empty(rendered.Fonts.Faces);
        Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.Contains(rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceContentTypeRejected);
        Assert.DoesNotContain(rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded);
    }

    [Fact]
    public void HtmlResourceSession_InlineAdmissionIsMimeAwareAtomicAndRecordsTheRequestedKind() {
        var session = new HtmlResourceSession(
            maxResourceBytes: 1024,
            maxTotalResourceBytes: 1024,
            maxResourceCount: 1,
            maxResourceRequests: 1);

        Assert.False(session.TryAcceptInline(HtmlResourceKind.Font, "data:image/png;base64,AA==",
            new HtmlResolvedResource(new byte[] { 0 }, "image/png"), out string rejectedCode, out _));
        Assert.Equal(HtmlRenderDiagnosticCodes.ResourceContentTypeRejected, rejectedCode);
        Assert.Equal(0, session.AcceptedResourceCount);
        Assert.Equal(0, session.AcceptedResourceBytes);

        Assert.True(session.TryAcceptInline(HtmlResourceKind.Font, "data:font/ttf;base64,AQ==",
            new HtmlResolvedResource(new byte[] { 1 }, "font/ttf"), out string acceptedCode, out _));
        Assert.Equal(string.Empty, acceptedCode);
        Assert.Equal(1, session.AcceptedResourceCount);
        Assert.Equal(1, session.AcceptedResourceBytes);
        Assert.Equal(HtmlResourceKind.Font, Assert.Single(session.Resources).Kind);
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
