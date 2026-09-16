using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeReactBuildApplicationTests {
    [Fact]
    public async Task ProductionBundleRoutesLoadsSplitModuleAndCapturesIndependentReport() {
        string directory = Path.Combine(AppContext.BaseDirectory, "Fixtures", "ReactBuild");
        string Read(string name) => File.ReadAllText(Path.Combine(directory, name));
        Uri origin = new("https://built-app.officeimo.test/index.html");
        var resources = new List<HtmlRuntimeResource> {
            HtmlRuntimeResource.FromText(new Uri(origin, "style.css"), Read("style.css"), "text/css"),
            HtmlRuntimeResource.FromText(new Uri(origin, "data.json"), Read("data.json"), "application/json")
        };
        resources.AddRange(Directory.EnumerateFiles(Path.Combine(directory, "dist"), "*.js")
            .Select(path => HtmlRuntimeResource.FromText(new Uri(origin, Path.GetFileName(path)), File.ReadAllText(path), "text/javascript")));
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);

        await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = origin,
            Html = Read("index.html"),
            Resources = resources,
            ReadyExpression = "document.querySelector('#total')?.textContent==='Total: 42'",
            Timeout = TimeSpan.FromSeconds(30)
        });
        await session.Locator("#total").WaitForTextAsync("Total: 42");
        var first = await session.CaptureAsync();

        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Region")).SelectOptionsAsync(new[] { "South" });
        await session.Locator("#total").WaitForTextAsync("Total: 18");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Report title")).FillAsync("Quarterly");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Add adjustment")).ClickAsync();
        await session.Locator("#total").WaitForTextAsync("Total: 21");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Prepare review")).ClickAsync();
        await session.Locator("h1").WaitForTextAsync("Review report");
        var review = await session.CaptureAsync("document.querySelector('h1')?.textContent==='Review report'");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Back to report")).ClickAsync();
        await session.Locator("h1").WaitForTextAsync("Regional report");
        Assert.Equal(origin, (await session.CaptureAsync("document.querySelector('#total')?.textContent==='Total: 21'")).DocumentUrl);
        await session.DisposeAsync();

        Assert.Equal("https://built-app.officeimo.test/review", review.DocumentUrl.AbsoluteUri);
        Assert.Equal("Regional report", first.Document.QuerySelector("h1")!.TextContent);
        Assert.Equal("Quarterly / South", review.Document.QuerySelector("#report-heading")!.TextContent);
        Assert.Equal("Total: 21", review.Document.QuerySelector("#total")!.TextContent);
        Assert.Contains(review.Resources, resource => resource.Url.AbsolutePath.Contains("chunk-review-", StringComparison.Ordinal));

        var conversion = HtmlConversionDocument.FromDocument(review.CreateStandaloneDocument(), new() { BaseUri = review.BaseUri });
        var retained = review.Resources.ToDictionary(resource => resource.Url.AbsoluteUri, StringComparer.Ordinal);
        HtmlRenderResourceResolver resolver = (request, _) => Task.FromResult(retained.TryGetValue(request.Uri.AbsoluteUri, out var resource)
            ? new HtmlResolvedResource(resource.Content, resource.ContentType, resource.FinalUrl, resource.RedirectCount) : null);
        var options = new HtmlToPdfOptions {
            ViewportWidth = 816,
            Margins = HtmlRenderMargins.All(0),
            ResourceResolver = resolver,
            ResourcePolicy = new PdfResourcePolicy {
                AllowSystemFontEmbedding = true,
                AllowDataUris = true,
                AllowEmbeddedPackageResources = true,
                AllowRemoteResourceResolution = true
            }
        };
        var screen = await HtmlRenderEngine.ExecuteAsync(conversion,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, options,
                HtmlRenderDocumentState.RuntimeSnapshot));
        byte[] screenPng = Assert.Single(screen.ExportImages()).Bytes;
        Assert.True(OfficePngReader.TryDecode(screenPng, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.FromRgb(243, 246, 250), image!.GetPixel(0, 0));
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_REACT_BUILD_EVIDENCE_DIR");
        if (!string.IsNullOrEmpty(evidence)) {
            Directory.CreateDirectory(evidence);
            File.WriteAllBytes(Path.Combine(evidence, "ScreenFullPage.png"), screenPng);
        }
        foreach (var intent in new[] { HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenSnapshotPaged }) {
            var result = await conversion.RenderToPdfResultAsync(HtmlRenderRequest.Create(intent, HtmlRenderEncoder.Pdf, options,
                HtmlRenderDocumentState.RuntimeSnapshot));
            byte[] bytes = result.ToBytes();
            var pdf = PdfReadDocument.Open(bytes);
            Assert.Contains("Quarterly / South", pdf.ExtractText());
            Assert.Contains("Total: 21", pdf.ExtractText());
            Assert.True(pdf.HasTaggedContent);
            if (!string.IsNullOrEmpty(evidence)) File.WriteAllBytes(Path.Combine(evidence, intent + ".pdf"), bytes);
        }
    }
}
