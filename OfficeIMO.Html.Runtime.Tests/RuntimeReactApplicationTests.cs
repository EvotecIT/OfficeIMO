using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeReactApplicationTests {
    [Fact]
    public async Task React18ReportLoadsAndCapturesFetchedData() {
        string directory = Path.Combine(AppContext.BaseDirectory, "Fixtures", "React18");
        string Read(string name) => File.ReadAllText(Path.Combine(directory, name));
        Uri origin = new("https://react.officeimo.test/index.html");
        HtmlRuntimeResource Resource(string name, string mediaType) =>
            HtmlRuntimeResource.FromText(new Uri(origin, name), Read(name), mediaType);
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);
        await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = origin,
            Html = Read("index.html"),
            Resources = new[] {
                Resource("react.production.min.js", "text/javascript"),
                Resource("react-dom.production.min.js", "text/javascript"),
                Resource("app.js", "text/javascript"),
                Resource("style.css", "text/css"),
                Resource("data.json", "application/json")
            }
        });
        await session.Locator("#total").WaitForTextAsync("Total: 42");
        var initial = await session.CaptureAsync();
        Assert.Equal("Regional report", initial.Document.QuerySelector("h1")!.TextContent);
        Assert.Equal("Total: 42", initial.Document.QuerySelector("[role=status]")!.TextContent);
        Assert.Equal("Monthly", initial.Document.QuerySelector("#report-title")!.FormState!.Value);

        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Region")).SelectOptionsAsync(new[] { "South" });
        await session.Locator("#total").WaitForTextAsync("Total: 18");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Add adjustment")).ClickAsync();
        await session.Locator("#total").WaitForTextAsync("Total: 21");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Report title")).FillAsync("Quarterly");
        await session.Locator("#report-heading").WaitForTextAsync("Quarterly / South");
        var edited = await session.CaptureAsync();
        Assert.Equal("Quarterly", edited.Document.QuerySelector("#report-title")!.FormState!.Value);
        Assert.Equal("Total: 21", edited.Document.QuerySelector("[aria-live=polite]")!.TextContent);
        Assert.Single(edited.Document.QuerySelectorAll("tbody tr"));

        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Prepare review")).ClickAsync();
        await session.Locator("h1").WaitForTextAsync("Review report");
        var review = await session.CaptureAsync();
        Assert.Null(review.Document.QuerySelector(".controls"));
        await session.DisposeAsync();

        Assert.Equal("Total: 42", initial.Document.QuerySelector("#total")!.TextContent);
        Assert.Equal("Monthly", initial.Document.QuerySelector("#report-title")!.FormState!.Value);
        Assert.Equal("Total: 21", review.Document.QuerySelector("#total")!.TextContent);

        var conversion = HtmlConversionDocument.FromDocument(review.CreateStandaloneDocument(), new() { BaseUri = review.BaseUri });
        var resources = review.Resources.ToDictionary(resource => resource.Url.AbsoluteUri, StringComparer.Ordinal);
        HtmlRenderResourceResolver resolver = (request, _) => Task.FromResult(resources.TryGetValue(request.Uri.AbsoluteUri, out var resource)
            ? new HtmlResolvedResource(resource.Content, resource.ContentType, resource.FinalUrl, resource.RedirectCount) : null);
        var options = new HtmlToPdfOptions {
            ViewportWidth = 816,
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
        byte[] png = Assert.Single(screen.ExportImages()).Bytes;
        Assert.Equal(new byte[] { 137, 80, 78, 71 }, png.Take(4));
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? screenImage));
        Assert.Equal(OfficeColor.FromRgb(243, 246, 250), screenImage!.GetPixel(0, 0));

        foreach (var profile in new[] { HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenSnapshotPaged }) {
            var raster = await HtmlRenderEngine.ExecuteAsync(conversion, HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Png, options,
                HtmlRenderDocumentState.RuntimeSnapshot));
            byte[] rasterPng = Assert.Single(raster.ExportImages()).Bytes;
            var result = await conversion.RenderToPdfResultAsync(HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options,
                HtmlRenderDocumentState.RuntimeSnapshot));
            byte[] pdf = result.ToBytes();
            Assert.NotEmpty(result.RenderResult.Document.Pages);
            string text = PdfReadDocument.Open(pdf).ExtractText();
            Assert.Contains("Review report", text);
            Assert.Contains("Total: 21", text);
            Assert.True(PdfReadDocument.Open(pdf).HasTaggedContent);
            if (Environment.GetEnvironmentVariable("OFFICEIMO_REACT_EVIDENCE_DIR") is { Length: > 0 } folder) {
                Directory.CreateDirectory(folder);
                File.WriteAllBytes(Path.Combine(folder, profile + ".pdf"), pdf);
                File.WriteAllBytes(Path.Combine(folder, profile + ".png"), rasterPng);
            }
        }
        if (Environment.GetEnvironmentVariable("OFFICEIMO_REACT_EVIDENCE_DIR") is { Length: > 0 } evidenceFolder) {
            Directory.CreateDirectory(evidenceFolder);
            File.WriteAllBytes(Path.Combine(evidenceFolder, "ScreenFullPage.png"), png);
        }
    }
}
