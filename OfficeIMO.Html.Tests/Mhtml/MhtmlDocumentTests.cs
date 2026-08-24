using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Tests;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class MhtmlDocumentTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RedirectTargetsPassTheFullUrlPolicyBeforeFetcherInvocation(bool rewriteTarget) {
        var document = new MhtmlDocument(
            "<img src='https://example.test/start.png'>",
            contentLocation: "https://example.test/page.html");
        var requested = new List<Uri>();
        var options = new HtmlRenderOptions {
            ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };
        options.ResourceUrlPolicy.ResolvedUrlTransform = value =>
            value.Contains("/blocked.png", StringComparison.Ordinal)
                ? rewriteTarget
                    ? value.Replace("/blocked.png", "/rewritten.png")
                    : null
                : value;
        MhtmlRemoteResourcePolicy policy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1);
        policy.ResourceFetcher = (request, cancellationToken) => {
            requested.Add(request.Uri);
            return Task.FromResult<MhtmlRemoteResourceResponse?>(
                MhtmlRemoteResourceResponse.Redirect(new Uri("/blocked.png", UriKind.Relative)));
        };
        document.ConfigureRenderOptions(options, policy);

        HtmlResolvedResource? resource = await options.ResourceResolver!(
            new HtmlRenderResourceRequest(
                new Uri("https://example.test/start.png"),
                "https://example.test/start.png",
                HtmlResourceKind.Image),
            CancellationToken.None);

        Assert.Null(resource);
        Assert.Equal(new[] { new Uri("https://example.test/start.png") }, requested);
    }

    [Fact]
    public void EmbeddedOnlyConfigurationPreservesCallerSharedResourceLimits() {
        var document = new MhtmlDocument("<p>Embedded only</p>");
        var options = new HtmlRenderOptions {
            MaxResourceBytes = 25L * 1024L * 1024L,
            MaxTotalResourceBytes = 100L * 1024L * 1024L,
            MaxResourceCount = 500,
            MaxResourceRequests = 1000,
            ResourceTimeout = TimeSpan.FromMinutes(2D)
        };

        document.ConfigureRenderOptions(options);

        Assert.Equal(25L * 1024L * 1024L, options.MaxResourceBytes);
        Assert.Equal(100L * 1024L * 1024L, options.MaxTotalResourceBytes);
        Assert.Equal(500, options.MaxResourceCount);
        Assert.Equal(1000, options.MaxResourceRequests);
        Assert.Equal(TimeSpan.FromMinutes(2D), options.ResourceTimeout);
    }

    [Fact]
    public async Task RedirectedStylesheetUsesFinalUriForDependenciesAndCanonicalIdentity() {
        var document = new MhtmlDocument(
            "<link rel='stylesheet' href='https://example.test/style.css'><p>Redirected style</p>",
            contentLocation: "https://example.test/page.html");
        var requested = new List<Uri>();
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        MhtmlRemoteResourcePolicy policy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1);
        policy.ResourceFetcher = (request, cancellationToken) => {
            requested.Add(request.Uri);
            MhtmlRemoteResourceResponse response = request.Uri.AbsolutePath switch {
                "/style.css" => MhtmlRemoteResourceResponse.Redirect(
                    new Uri("/assets/style.css", UriKind.Relative)),
                "/assets/style.css" => new MhtmlRemoteResourceResponse(
                    Encoding.UTF8.GetBytes("p{background-image:url(image.png)}"), "text/css"),
                "/assets/image.png" => new MhtmlRemoteResourceResponse(png, "image/png"),
                _ => new MhtmlRemoteResourceResponse(Array.Empty<byte>(), "application/octet-stream")
            };
            return Task.FromResult<MhtmlRemoteResourceResponse?>(response);
        };
        var options = new HtmlRenderOptions();
        document.ConfigureRenderOptions(options, policy);

        HtmlResourceSession session = await HtmlResourceSession.ResolveAsync(
            document.HtmlDocument.ResourceManifest, options);

        Assert.Contains(new Uri("https://example.test/assets/image.png"), requested);
        Assert.DoesNotContain(new Uri("https://example.test/image.png"), requested);
        Assert.Contains(session.Resources, entry =>
            entry.Kind == HtmlResourceKind.Stylesheet
            && entry.CanonicalSource == "https://example.test/assets/style.css");
    }

    [Fact]
    public void RemotePolicyRejectsRelationallyInconsistentResourceLimits() {
        var document = new MhtmlDocument("<p>Limits</p>", contentLocation: "https://example.test/page.html");
        MhtmlRemoteResourcePolicy bytePolicy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile();
        bytePolicy.MaximumResourceBytes = 10;
        bytePolicy.MaximumTotalResourceBytes = 5;

        ArgumentOutOfRangeException byteException = Assert.Throws<ArgumentOutOfRangeException>(
            () => document.ConfigureRenderOptions(new HtmlRenderOptions(), bytePolicy));
        Assert.Equal(nameof(MhtmlRemoteResourcePolicy.MaximumTotalResourceBytes), byteException.ParamName);

        MhtmlRemoteResourcePolicy requestPolicy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile();
        requestPolicy.MaximumResourceCount = 2;
        requestPolicy.MaximumResourceRequests = 1;

        ArgumentOutOfRangeException requestException = Assert.Throws<ArgumentOutOfRangeException>(
            () => document.ConfigureRenderOptions(new HtmlRenderOptions(), requestPolicy));
        Assert.Equal(nameof(MhtmlRemoteResourcePolicy.MaximumResourceRequests), requestException.ParamName);
    }

    [Fact]
    public async Task ConcurrentRedirectBudgetFailuresEmitOneDiagnosticPerResource() {
        var document = new MhtmlDocument(
            "<img src='https://example.test/a.png'><img src='https://example.test/b.png'>",
            contentLocation: "https://example.test/page.html");
        int calls = 0;
        var bothFetchesStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var options = new HtmlRenderOptions();
        MhtmlRemoteResourcePolicy policy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1);
        policy.MaximumResourceCount = 2;
        policy.MaximumResourceRequests = 2;
        policy.ResourceFetcher = async (request, cancellationToken) => {
            if (Interlocked.Increment(ref calls) == 2) bothFetchesStarted.TrySetResult(true);
            await bothFetchesStarted.Task.ConfigureAwait(false);
            return MhtmlRemoteResourceResponse.Redirect(new Uri(request.Uri.AbsolutePath + ".next", UriKind.Relative));
        };
        document.ConfigureRenderOptions(options, policy);

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document.HtmlDocument, options);

        Assert.Equal(2, calls);
        Assert.Equal(2, rendered.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded));
    }

    [Fact]
    public async Task RedirectHopsConsumeTheOperationWideResourceRequestBudget() {
        var document = new MhtmlDocument("<img src='https://example.test/start.png'>",
            contentLocation: "https://example.test/page.html");
        int calls = 0;
        var options = new HtmlRenderOptions();
        MhtmlRemoteResourcePolicy policy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 2);
        policy.MaximumResourceCount = 1;
        policy.MaximumResourceRequests = 1;
        policy.ResourceFetcher = (request, cancellationToken) => {
            calls++;
            return Task.FromResult<MhtmlRemoteResourceResponse?>(
                MhtmlRemoteResourceResponse.Redirect(new Uri("/next.png", UriKind.Relative)));
        };
        document.ConfigureRenderOptions(options, policy);

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document.HtmlDocument, options);

        Assert.Equal(1, calls);
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded);
    }

    [Fact]
    public void ConfigureRenderOptionsRetainsTheOriginalBinarySignature() {
        Assert.NotNull(typeof(MhtmlDocument).GetMethod(
            nameof(MhtmlDocument.ConfigureRenderOptions),
            new[] { typeof(HtmlRenderOptions) }));
    }

    [Fact]
    public async Task LoadSelectsDeclaredRootAndResolvesCidResource() {
        const string archive = "MIME-Version: 1.0\r\n" +
            "Subject: Saved page\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"; start=\"<root>\"\r\n\r\n" +
            "--archive\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQID\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n" +
            "Content-Location: https://example.test/page/index.html\r\n\r\n" +
            "<html><body><img src=\"cid:logo\"></body></html>\r\n" +
            "--archive--\r\n";
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(archive));

        MhtmlDocument document = MhtmlDocument.Load(stream);
        HtmlResolvedResource? resolved = await document.CreateResourceResolver()(
            new HtmlRenderResourceRequest(new Uri("cid:logo"), "cid:logo", HtmlResourceKind.Image),
            CancellationToken.None);

        Assert.Contains("cid:logo", document.Html, StringComparison.Ordinal);
        Assert.Equal("root", document.RootContentId);
        Assert.Equal("https://example.test/page/index.html", document.ContentLocation);
        Assert.Equal(new Uri("https://example.test/page/index.html"), document.BaseUri);
        MhtmlResource resource = Assert.Single(document.Resources);
        Assert.Equal("logo", resource.ContentId);
        Assert.NotNull(resolved);
        Assert.Equal(new byte[] { 1, 2, 3 }, resolved!.Bytes);
        Assert.Equal("image/png", resolved.ContentType);
    }

    [Fact]
    public void LoadDoesNotExposeFirstRelatedHtmlRootAsAResource() {
        const string archive =
            "MIME-Version: 1.0\r\n" +
            "Subject: root first\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"; start=\"<root>\"\r\n\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n" +
            "Content-Location: https://example.test/page/index.html\r\n\r\n" +
            "<html><body><img src=\"cid:logo\"></body></html>\r\n" +
            "--archive\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQID\r\n" +
            "--archive--\r\n";
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(archive));

        MhtmlDocument document = MhtmlDocument.Load(stream);

        MhtmlResource resource = Assert.Single(document.Resources);
        Assert.Equal("logo", resource.ContentId);
        Assert.Equal(new byte[] { 1, 2, 3 }, resource.Content);
    }

    [Fact]
    public void ResourceContentRemainsAnImmutableSnapshotAfterConstructionAndLoad() {
        byte[] input = { 1, 2, 3 };
        var resource = new MhtmlResource(input, "image/png", contentId: "logo");
        input[0] = 9;
        byte[] firstRead = resource.Content;
        firstRead[1] = 9;

        var source = new MhtmlDocument("<img src='cid:logo'>", new[] { resource }, rootContentId: "root");
        using var stream = new MemoryStream(source.ToBytes());
        MhtmlResource loaded = Assert.Single(MhtmlDocument.Load(stream).Resources);
        byte[] loadedRead = loaded.Content;
        loadedRead[2] = 9;

        Assert.Equal(new byte[] { 1, 2, 3 }, resource.Content);
        Assert.Equal(new byte[] { 1, 2, 3 }, loaded.Content);
    }

    [Fact]
    public async Task ConfigureRenderOptionsAllowsPackageResourcesWithoutRelaxingHyperlinksOrFallbacks() {
        int fallbackCalls = 0;
        var document = new MhtmlDocument(
            "<a href='cid:logo'>link</a><img src='cid:logo'>",
            new[] { new MhtmlResource(new byte[] { 1, 2, 3 }, "image/png", contentId: "logo") },
            "file:///snapshot/page.html");
        var options = new HtmlRenderOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceResolver = (request, cancellationToken) => {
                fallbackCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(new byte[] { 4, 5, 6 }, "image/png"));
            }
        };

        document.ConfigureRenderOptions(options);

        Assert.DoesNotContain("cid", options.UrlPolicy.AllowedUrlSchemes);
        Assert.DoesNotContain(Uri.UriSchemeFile, options.UrlPolicy.AllowedUrlSchemes);
        Assert.True(options.UrlPolicy.DisallowFileUrls);
        Assert.NotNull(options.ResourceUrlPolicy);
        Assert.Contains("cid", options.ResourceUrlPolicy!.AllowedUrlSchemes);
        Assert.Contains(Uri.UriSchemeFile, options.ResourceUrlPolicy.AllowedUrlSchemes);
        Assert.False(options.ResourceUrlPolicy.DisallowFileUrls);
        Assert.NotNull(options.ResourceResolver);
        HtmlResolvedResource? embedded = await options.ResourceResolver!(
            new HtmlRenderResourceRequest(new Uri("cid:logo"), "cid:logo", HtmlResourceKind.Image),
            CancellationToken.None);
        HtmlResolvedResource? missingFile = await options.ResourceResolver(
            new HtmlRenderResourceRequest(new Uri("file:///outside/secret.png"), "file:///outside/secret.png", HtmlResourceKind.Image),
            CancellationToken.None);
        Assert.NotNull(embedded);
        Assert.Null(missingFile);
        Assert.Equal(0, fallbackCalls);
    }

    [Fact]
    public void ConversionDocumentPreservesOnlyArchiveBackedCidAndFileResources() {
        var document = new MhtmlDocument(
            "<a href='cid:logo'>package link</a><img src='cid:logo'><img src='images/chart.png'><img src='file:///outside/secret.png'>",
            new[] {
                new MhtmlResource(new byte[] { 1 }, "image/png", contentId: "logo", fileName: "logo.png"),
                new MhtmlResource(new byte[] { 2 }, "image/png", contentLocation: "images/chart.png", fileName: "chart.png")
            },
            "file:///snapshot/page.html");

        string html = document.HtmlDocument.HtmlForConversion;

        Assert.Contains("src=\"cid:logo\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("file:///snapshot/images/chart.png", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("file:///outside/secret.png", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=\"cid:logo\"", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConstructedArchiveRoundTripsUnreferencedRelatedResource() {
        var resource = new MhtmlResource(Encoding.UTF8.GetBytes("body { color: black; }"),
            "text/css", contentLocation: "styles/site.css", fileName: "site.css");
        var document = new MhtmlDocument("<html><body>saved</body></html>", new[] { resource },
            "https://example.test/page/index.html", "root", "Saved page");

        byte[] bytes = document.ToBytes();
        string serialized = Encoding.ASCII.GetString(bytes);
        using var stream = new MemoryStream(bytes);
        MhtmlDocument roundTrip = MhtmlDocument.Load(stream);

        Assert.Contains("multipart/related", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("type=\"text/html\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("start=\"<root>\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Saved page", roundTrip.Subject);
        Assert.Equal("root", roundTrip.RootContentId);
        Assert.Equal("styles/site.css", Assert.Single(roundTrip.Resources).ContentLocation);
    }

    [Fact]
    public void ConstructedArchiveWithoutResourcesStillUsesRelatedContainer() {
        var document = new MhtmlDocument("<html><body>standalone</body></html>", rootContentId: "root");

        byte[] bytes = document.ToBytes();
        using var stream = new MemoryStream(bytes);
        MhtmlDocument roundTrip = MhtmlDocument.Load(stream);

        Assert.Equal("root", roundTrip.RootContentId);
        Assert.Empty(roundTrip.Resources);
    }

    [Fact]
    public void LoadRejectsOrdinaryEmailMessage() {
        const string message = "Subject: ordinary\r\nContent-Type: text/html; charset=utf-8\r\n\r\n<p>mail</p>\r\n";
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(message));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => MhtmlDocument.Load(stream));

        Assert.Contains("multipart/related", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DuplicateIdentifiersAndLocationsAreFirstWinsAndDiagnosed() {
        var document = new MhtmlDocument(
            "<img src='cid:logo'><img src='images/logo.png'>",
            new[] {
                new MhtmlResource(new byte[] { 1 }, "image/png", "logo", "images/logo.png"),
                new MhtmlResource(new byte[] { 2 }, "image/png", "LOGO", "./images/logo.png")
            },
            "https://example.test/archive/page.html");

        HtmlRenderResourceResolver resolver = document.CreateResourceResolver();
        HtmlResolvedResource? byId = await resolver(
            new HtmlRenderResourceRequest(new Uri("cid:logo"), "cid:logo", HtmlResourceKind.Image),
            CancellationToken.None);
        HtmlResolvedResource? byLocation = await resolver(
            new HtmlRenderResourceRequest(new Uri("https://example.test/archive/images/logo.png"), "images/logo.png", HtmlResourceKind.Image),
            CancellationToken.None);

        Assert.Equal(new byte[] { 1 }, byId!.Bytes);
        Assert.Equal(new byte[] { 1 }, byLocation!.Bytes);
        Assert.Contains(document.MimeDiagnostics, diagnostic => diagnostic.Code == MhtmlDiagnosticCodes.DuplicateContentId);
        Assert.Contains(document.MimeDiagnostics, diagnostic => diagnostic.Code == MhtmlDiagnosticCodes.DuplicateContentLocation);
    }

    [Fact]
    public void MalformedClosingBoundaryIsRecoveredWithStableMimeDiagnostic() {
        const string archive = "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"\r\n\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\n\r\n" +
            "<p>Recovered</p>\r\n";
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(archive));

        MhtmlDocument document = MhtmlDocument.Load(stream);

        Assert.Contains("Recovered", document.Html, StringComparison.Ordinal);
        Assert.Contains(document.MimeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_BOUNDARY_NOT_CLOSED");
    }

    [Fact]
    public void LegacyCharsetAndNestedRelatedResourcesFlowThroughSharedMimeReader() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding legacy = Encoding.GetEncoding(1250);
        const string header = "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=outer; type=\"text/html\"\r\n\r\n" +
            "--outer\r\nContent-Type: multipart/alternative; boundary=inner\r\n\r\n" +
            "--inner\r\nContent-Type: text/plain; charset=windows-1250\r\n\r\nFallback\r\n" +
            "--inner\r\nContent-Type: text/html; charset=windows-1250\r\n" +
            "Content-Location: https://example.test/page.html\r\n\r\n";
        const string tail = "\r\n--inner--\r\n" +
            "--outer\r\nContent-Type: image/png\r\nContent-ID: <nested-logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQID\r\n--outer--\r\n";
        byte[] prefix = Encoding.ASCII.GetBytes(header);
        byte[] html = legacy.GetBytes("<p>Zażółć gęślą</p><img src='cid:nested-logo'>");
        byte[] suffix = Encoding.ASCII.GetBytes(tail);
        using var stream = new MemoryStream(prefix.Concat(html).Concat(suffix).ToArray());

        MhtmlDocument document = MhtmlDocument.Load(stream);

        Assert.Contains("Zażółć gęślą", document.Html, StringComparison.Ordinal);
        Assert.Equal("nested-logo", Assert.Single(document.Resources).ContentId);
    }

    [Fact]
    public async Task RemoteFallbackIsOfflineByDefaultAndSameOriginRedirectBoundedWhenEnabled() {
        var document = new MhtmlDocument("<img src='https://example.test/missing.png'>",
            contentLocation: "https://example.test/page.html");
        int offlineCalls = 0;
        var offline = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => {
                offlineCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(new byte[] { 1 }, "image/png"));
            }
        };
        document.ConfigureRenderOptions(offline);

        HtmlResolvedResource? offlineResult = await offline.ResourceResolver!(
            new HtmlRenderResourceRequest(new Uri("https://example.test/missing.png"), "https://example.test/missing.png", HtmlResourceKind.Image),
            CancellationToken.None);

        Assert.Null(offlineResult);
        Assert.Equal(0, offlineCalls);

        int boundedCalls = 0;
        var bounded = new HtmlRenderOptions();
        MhtmlRemoteResourcePolicy boundedPolicy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1);
        boundedPolicy.ResourceFetcher = (request, cancellationToken) => {
            boundedCalls++;
            return Task.FromResult<MhtmlRemoteResourceResponse?>(request.RedirectNumber == 0
                ? MhtmlRemoteResourceResponse.Redirect(new Uri("/final.png", UriKind.Relative))
                : new MhtmlRemoteResourceResponse(new byte[] { 2 }, "image/png"));
        };
        document.ConfigureRenderOptions(bounded, boundedPolicy);

        HtmlResolvedResource? accepted = await bounded.ResourceResolver!(
            new HtmlRenderResourceRequest(new Uri("https://example.test/missing.png"), "https://example.test/missing.png", HtmlResourceKind.Image),
            CancellationToken.None);
        Assert.NotNull(accepted);
        Assert.Equal(2, boundedCalls);
        Assert.Equal(new Uri("https://example.test/final.png"), accepted!.FinalUri);
        Assert.Equal(1, accepted.RedirectCount);

        var missingFetcher = new HtmlRenderOptions();
        document.ConfigureRenderOptions(missingFetcher, MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1));
        HtmlResolvedResource? missingFetcherResult = await missingFetcher.ResourceResolver!(
            new HtmlRenderResourceRequest(new Uri("https://example.test/missing.png"), "https://example.test/missing.png", HtmlResourceKind.Image),
            CancellationToken.None);
        Assert.Null(missingFetcherResult);

        int redirectedCalls = 0;
        var redirected = new HtmlRenderOptions();
        MhtmlRemoteResourcePolicy redirectedPolicy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1);
        redirectedPolicy.ResourceFetcher = (request, cancellationToken) => {
            redirectedCalls++;
            return Task.FromResult<MhtmlRemoteResourceResponse?>(
                MhtmlRemoteResourceResponse.Redirect(new Uri("https://other.test/image.png")));
        };
        document.ConfigureRenderOptions(redirected, redirectedPolicy);
        HtmlResolvedResource? rejected = await redirected.ResourceResolver!(
            new HtmlRenderResourceRequest(new Uri("https://example.test/missing.png"), "https://example.test/missing.png", HtmlResourceKind.Image),
            CancellationToken.None);

        Assert.Null(rejected);
        Assert.Equal(1, redirectedCalls);
    }

    [Fact]
    public void ActiveContentRemainsFilteredByManagedHtmlTier() {
        var document = new MhtmlDocument("<script>alert(1)</script><p>safe</p>");

        Assert.DoesNotContain("alert(1)", document.HtmlDocument.HtmlForConversion, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<script></script>", document.HtmlDocument.HtmlForConversion, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("safe", document.HtmlDocument.HtmlForConversion, StringComparison.Ordinal);
    }

    [Fact]
    public void MalformedLegacyNestedArchiveConvertsThroughManagedPdfTierWithMimeDiagnostics() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding legacy = Encoding.GetEncoding(1250);
        const string header = "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=outer; type=\"text/html\"\r\n\r\n" +
            "--outer\r\nContent-Type: multipart/alternative; boundary=inner\r\n\r\n" +
            "--inner\r\nContent-Type: text/html; charset=windows-1250\r\n" +
            "Content-Location: https://example.test/page.html\r\n\r\n";
        string tail = "\r\n--inner--\r\n" +
            "--outer\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\n" +
            Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2)) + "\r\n";
        byte[] prefix = Encoding.ASCII.GetBytes(header);
        byte[] html = legacy.GetBytes("<script>document.write('executed')</script><h1>Zażółć MHTML</h1><img src='cid:logo'>");
        byte[] suffix = Encoding.ASCII.GetBytes(tail);
        using var stream = new MemoryStream(prefix.Concat(html).Concat(suffix).ToArray());

        MhtmlDocument document = MhtmlDocument.Load(stream);
        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult();
        string text = PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText();

        Assert.Contains("Zażółć MHTML", text, StringComparison.Ordinal);
        Assert.DoesNotContain("executed", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Warnings, warning => warning.Code == "EMAIL_MIME_BOUNDARY_NOT_CLOSED");
    }

    [Fact]
    public async Task ExplicitSameOriginRemotePolicyFlowsThroughMhtmlPdfAndRejectsRedirectEscape() {
        var document = new MhtmlDocument(
            "<img src='https://example.test/allowed.png'><img src='https://example.test/escaped.png'>",
            contentLocation: "https://example.test/page.html");
        int calls = 0;
        var options = new HtmlPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
        };
        MhtmlRemoteResourcePolicy remotePolicy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 1);
        remotePolicy.ResourceFetcher = (request, cancellationToken) => {
                calls++;
                return Task.FromResult<MhtmlRemoteResourceResponse?>(
                    request.Uri.AbsolutePath.EndsWith("escaped.png", StringComparison.Ordinal)
                        ? MhtmlRemoteResourceResponse.Redirect(new Uri("https://other.test/escaped.png"))
                        : new MhtmlRemoteResourceResponse(PdfPngTestImages.CreateRgbPng(2, 2), "image/png"));
        };
        document.ConfigureRenderOptions(options, remotePolicy);

        PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options);

        Assert.Equal(2, calls);
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()), image => image.IsImageFile);
        Assert.Contains(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }
}
