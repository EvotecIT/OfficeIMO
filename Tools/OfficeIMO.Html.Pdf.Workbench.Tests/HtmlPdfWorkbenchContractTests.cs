using System.Text;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Pdf.Workbench.Tests;

public sealed class HtmlPdfWorkbenchContractTests {
    [Fact]
    public void Templates_AreUniqueBoundedAndParseable() {
        Assert.Equal(HtmlPdfWorkbenchTemplates.All.Count, HtmlPdfWorkbenchTemplates.All.Select(template => template.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.All(HtmlPdfWorkbenchTemplates.All, template => {
            Assert.False(string.IsNullOrWhiteSpace(template.Name));
            Assert.InRange(template.Html.Length + template.Css.Length, 1, HtmlPdfWorkbenchConversionService.MaximumInputCharacters);
            _ = HtmlConversionDocument.Parse(template.Html);
        });
    }

    [Fact]
    public void PreviewComposer_InjectsOfflinePolicyAndContainsStyleMarkup() {
        string preview = HtmlPdfPreviewComposer.Compose(
            "<!-- <head> --><meta http-equiv=\"refresh\" content=\"0;url=https://example.com\"><iframe src=\"https://example.com\"></iframe><img src=\"https://example.com/a.png\">",
            "body{color:red}</style><script>alert(1)</script>");

        Assert.Contains("Content-Security-Policy", preview, StringComparison.Ordinal);
        Assert.Contains("connect-src 'none'", preview, StringComparison.Ordinal);
        Assert.DoesNotContain("http-equiv=\"refresh\"", preview, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<iframe", preview, StringComparison.OrdinalIgnoreCase);
        Assert.True(
            preview.IndexOf("Content-Security-Policy", StringComparison.Ordinal) < preview.IndexOf("example.com/a.png", StringComparison.Ordinal),
            "The fixed preview policy must precede all untrusted markup.");
        Assert.Contains("<\\/style>", preview, StringComparison.Ordinal);
        Assert.DoesNotContain("body{color:red}</style><script>", preview, StringComparison.Ordinal);
        string capture = HtmlPdfPreviewComposer.ComposeForCapture("<p>Capture</p>", "p{color:blue}");
        Assert.DoesNotContain("Content-Security-Policy", capture, StringComparison.Ordinal);
        Assert.Contains("p{color:blue}", capture, StringComparison.Ordinal);

        string restrictedCapture = HtmlPdfPreviewComposer.ComposeForCapture(
            "<meta http-equiv=\"Content-Security-Policy\" content=\"style-src 'none'\"><p>Capture</p>",
            "p{color:green}");
        Assert.DoesNotContain("Content-Security-Policy", restrictedCapture, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("p{color:green}", restrictedCapture, StringComparison.Ordinal);

        string localizedCapture = HtmlPdfPreviewComposer.ComposeForCapture(
            "<!doctype html><html lang=\"en\"><body>Capture</body></html>",
            string.Empty,
            "pl-PL");
        Assert.Contains("lang=\"pl-PL\"", localizedCapture, StringComparison.Ordinal);
        Assert.DoesNotContain("lang=\"en\"", localizedCapture, StringComparison.Ordinal);
    }

    [Fact]
    public void InputFingerprint_LengthPrefixesHtmlAndCssFields() {
        string first = HtmlPdfWorkbenchConversionService.ComputeInputSha256("a\u001eb", "c");
        string second = HtmlPdfWorkbenchConversionService.ComputeInputSha256("a", "b\u001ec");

        Assert.NotEqual(first, second);
        Assert.Equal(first, HtmlPdfWorkbenchConversionService.ComputeInputSha256("a\u001eb", "c"));
    }

    [Theory]
    [InlineData("127.0.0.1", 5105, true)]
    [InlineData("localhost", 5105, false)]
    [InlineData("attacker.example", 5105, false)]
    [InlineData("127.0.0.1", 5106, false)]
    public void RequestBoundary_RequiresTheConfiguredHostAndPort(string host, int port, bool expected) {
        var listenUri = new Uri("http://127.0.0.1:5105");
        Assert.Equal(expected, WorkbenchRequestBoundary.IsAllowedHost(new Microsoft.AspNetCore.Http.HostString(host, port), listenUri));
    }

    [Theory]
    [InlineData("http://127.0.0.1:5105", true)]
    [InlineData("http://localhost:5105", false)]
    [InlineData("https://127.0.0.1:5105", false)]
    [InlineData("https://attacker.example", false)]
    public void RequestBoundary_RequiresAnExactWebSocketOrigin(string origin, bool expected) {
        Assert.Equal(expected, WorkbenchRequestBoundary.IsAllowedWebSocketOrigin(origin, new Uri("http://127.0.0.1:5105")));
    }

    [Fact]
    public void LaunchProfiles_MatchTheEnforcedLoopbackEndpoint() {
        using JsonDocument settings = JsonDocument.Parse(File.ReadAllText(
            Path.Combine(AppContext.BaseDirectory, "launchSettings.json")));
        JsonElement root = settings.RootElement;

        Assert.Equal(
            "http://127.0.0.1:5105",
            root.GetProperty("iisSettings").GetProperty("iisExpress").GetProperty("applicationUrl").GetString());
        Assert.Equal(
            "http://127.0.0.1:5105",
            root.GetProperty("profiles").GetProperty("http").GetProperty("applicationUrl").GetString());
    }

    [Fact]
    public async Task ManagedConversion_ProducesInspectablePdfAndMatchingEvidence() {
        await using var renderer = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(networkPolicy: HtmlBrowserNetworkPolicy.Offline));
        var service = new HtmlPdfWorkbenchConversionService(renderer);
        HtmlPdfWorkbenchTemplate template = HtmlPdfWorkbenchTemplates.Find("accessible-report");

        HtmlPdfWorkbenchResult result = await service.ConvertAsync(new HtmlPdfWorkbenchRequest(
            template.Html,
            template.Css,
            HtmlPdfWorkbenchEngine.Managed,
            new HtmlPdfWorkbenchSettings()));

        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(result.PdfBytes, 0, 8), StringComparison.Ordinal);
        Assert.Equal("Managed", result.Evidence.Engine);
        Assert.Equal(result.PdfBytes.Length, result.Evidence.PdfBytes);
        Assert.True(PdfDocument.Open(result.PdfBytes).Inspect().PageCount > 0);
        using JsonDocument evidence = JsonDocument.Parse(result.EvidenceBytes);
        Assert.Equal("officeimo.html-pdf-workbench/v1", evidence.RootElement.GetProperty("schema").GetString());
        Assert.Equal(result.Evidence.OutputSha256, evidence.RootElement.GetProperty("outputSha256").GetString());
    }

    [Fact]
    public async Task ConvertAsync_OversizedInputFailsBeforeRendering() {
        await using var renderer = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(networkPolicy: HtmlBrowserNetworkPolicy.Offline));
        var service = new HtmlPdfWorkbenchConversionService(renderer);
        string oversized = "<p>" + new string('x', HtmlPdfWorkbenchConversionService.MaximumInputCharacters) + "</p>";

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() => service.ConvertAsync(new HtmlPdfWorkbenchRequest(
            oversized,
            string.Empty,
            HtmlPdfWorkbenchEngine.Managed,
            new HtmlPdfWorkbenchSettings())));

        Assert.Contains("workbench limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ArtifactStore_ReturnsOpaqueExpiringLinksAndPayloads() {
        var store = new WorkbenchArtifactStore();
        var evidence = new HtmlPdfWorkbenchEvidence(
            "test", DateTimeOffset.UtcNow, "Managed", "test", "input", "output", 1, 4, 1, false,
            new HtmlPdfWorkbenchSettings(), Array.Empty<HtmlPdfWorkbenchDiagnostic>(), null);
        WorkbenchArtifactLink link = store.Add(new HtmlPdfWorkbenchResult(new byte[] { 1, 2, 3, 4 }, new byte[] { 5, 6 }, evidence));

        Assert.Equal(48, link.Token.Length);
        Assert.StartsWith("/workbench/artifacts/", link.PdfUrl, StringComparison.Ordinal);
        Assert.True(store.TryGet(link.Token, out WorkbenchArtifact? artifact));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, artifact!.PdfBytes);
        Assert.False(store.TryGet("not-a-token", out _));
    }

    [Fact]
    [Trait("Category", "Live")]
    public async Task ChromiumConversion_ProducesBrowserEvidenceUnderOfflinePolicy() {
        await using var renderer = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(
            maximumBrowserInstances: 1,
            maximumQueuedCaptures: 0,
            networkPolicy: HtmlBrowserNetworkPolicy.Offline));
        var service = new HtmlPdfWorkbenchConversionService(renderer);
        HtmlPdfWorkbenchTemplate template = HtmlPdfWorkbenchTemplates.Default;

        HtmlPdfWorkbenchResult result = await service.ConvertAsync(new HtmlPdfWorkbenchRequest(
            template.Html,
            template.Css,
            HtmlPdfWorkbenchEngine.Chromium,
            new HtmlPdfWorkbenchSettings()));

        Assert.Equal("Chromium", result.Evidence.Engine);
        Assert.NotNull(result.Evidence.Browser);
        Assert.False(string.IsNullOrWhiteSpace(result.Evidence.Browser!.BrowserVersion));
        Assert.Equal(0, result.Evidence.Browser.BlockedRequestCount);
        Assert.True(PdfDocument.Open(result.PdfBytes).Inspect().PageCount > 0);
    }

    [Fact]
    [Trait("Category", "Live")]
    public async Task ChromiumConversion_AppliesConfiguredDocumentLanguage() {
        await using var renderer = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(
            maximumBrowserInstances: 1,
            maximumQueuedCaptures: 0,
            networkPolicy: HtmlBrowserNetworkPolicy.Offline));
        var service = new HtmlPdfWorkbenchConversionService(renderer);
        var settings = new HtmlPdfWorkbenchSettings {
            Language = "pl-PL",
            TaggedPdf = true
        };

        HtmlPdfWorkbenchResult result = await service.ConvertAsync(new HtmlPdfWorkbenchRequest(
            "<!doctype html><html lang=\"en\"><body><p>Język polski</p></body></html>",
            string.Empty,
            HtmlPdfWorkbenchEngine.Chromium,
            settings));

        Assert.Equal("pl-PL", PdfReadDocument.Open(result.PdfBytes).CatalogLanguage);
    }

    [Fact]
    [Trait("Category", "Live")]
    public async Task ChromiumConversion_ReportsResourcesBlockedByOfflinePolicy() {
        await using var renderer = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(
            maximumBrowserInstances: 1,
            maximumQueuedCaptures: 0,
            networkPolicy: HtmlBrowserNetworkPolicy.Offline));
        var service = new HtmlPdfWorkbenchConversionService(renderer);

        HtmlPdfWorkbenchResult result = await service.ConvertAsync(new HtmlPdfWorkbenchRequest(
            "<!doctype html><html><body><h1>Offline proof</h1><img src=\"https://example.com/remote.png\" alt=\"remote\"></body></html>",
            "h1{color:#123456}",
            HtmlPdfWorkbenchEngine.Chromium,
            new HtmlPdfWorkbenchSettings()));

        Assert.NotNull(result.Evidence.Browser);
        Assert.True(result.Evidence.Browser!.BlockedRequestCount > 0);
        Assert.True(result.Evidence.HasLoss);
        Assert.Contains(result.Evidence.Diagnostics, diagnostic => diagnostic.Code == "BrowserRequestBlocked");
    }
}
