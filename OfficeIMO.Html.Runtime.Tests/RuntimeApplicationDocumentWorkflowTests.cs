using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Rendering;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeApplicationDocumentWorkflowTests {
    [Theory]
    [InlineData("vanilla")]
    [InlineData("react-build")]
    [InlineData("preact")]
    [InlineData("legacy")]
    public async Task OneWorkflowCapturesAndRendersFourApplicationClasses(string caseId) {
        (HtmlScriptRequest page, HtmlAutomationRequest[] actions, string ready, string expectedText) = CreateCase(caseId);
        IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);
        var rendering = new HtmlToPdfOptions { ViewportWidth = 816D, Margins = HtmlRenderMargins.All(0D) };
        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                Actions = actions,
                FinalReadyExpression = ready,
                RenderRequests = new[] {
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, rendering),
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, rendering),
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, rendering)
                }
            });

        Assert.Equal("officeimo.trusted-process", result.Provider.Id);
        Assert.True(result.Capture.Document.IsReadOnly);
        Assert.Equal(actions.Length, result.Actions.Count);
        Assert.All(result.Actions, action => Assert.Equal(HtmlAutomationStatus.Success, action.Status));
        Assert.Contains(result.Trace.Events, item => item.Kind == HtmlRuntimeEventKind.Capture);
        Assert.Equal(3, result.Outputs.Count);
        Assert.All(result.Outputs, output => Assert.Equal(HtmlRenderDocumentState.RuntimeSnapshot, output.Render.Request.DocumentState));
        byte[] png = Assert.Single(result.Outputs[0].Images).Bytes;
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Null(result.Outputs[0].Pdf);
        foreach (HtmlApplicationRenderOutput output in result.Outputs.Skip(1)) {
            Assert.Empty(output.Images);
            Assert.NotNull(output.Pdf);
            PdfReadDocument pdf = PdfReadDocument.Open(output.Pdf!.ToBytes());
            Assert.Contains(expectedText, pdf.ExtractText(), StringComparison.Ordinal);
            Assert.True(pdf.HasTaggedContent);
        }
        if (caseId == "react-build") {
            Assert.Equal("https://built-app.officeimo.test/review", result.Capture.DocumentUrl.AbsoluteUri);
            Assert.Contains(result.Capture.Resources, resource => resource.Url.AbsolutePath.Contains("chunk-review-", StringComparison.Ordinal));
            Assert.Equal(OfficeColor.FromRgb(243, 246, 250), image!.GetPixel(0, 0));
        }
        if (Environment.GetEnvironmentVariable("OFFICEIMO_APPLICATION_EVIDENCE_DIR") is { Length: > 0 } evidenceRoot) {
            string folder = Path.Combine(evidenceRoot, caseId);
            Directory.CreateDirectory(folder);
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen.png"), png);
            File.WriteAllBytes(Path.Combine(folder, "officeimo-print.pdf"), result.Outputs[1].Pdf!.ToBytes());
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen-to-page.pdf"), result.Outputs[2].Pdf!.ToBytes());
        }
    }

    [Fact]
    public async Task ExternalRenderResolverIsRejectedBeforeStartingAContext() {
        var host = new CountingHost();
        var options = new HtmlRenderOptions {
            ResourceResolver = (_, _) => Task.FromResult<HtmlResolvedResource?>(null)
        };
        var request = new HtmlApplicationDocumentRequest {
            RenderRequests = new[] { HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, options) }
        };

        await Assert.ThrowsAsync<ArgumentException>(() => HtmlApplicationDocumentWorkflow.RunAsync(host, request));
        Assert.Equal(0, host.ContextsCreated);
    }

    [Fact]
    public async Task SuppliedRenderOnlyImageRemainsAvailableAfterTheLivePageCloses() {
        Uri imageUrl = new("https://legacy.officeimo.test/print-background.png");
        byte[] png = OfficePngWriter.EncodeRgba(1, 1, new byte[] { 255, 0, 0, 255 });
        HtmlScriptRequest page = LegacyCase().Item1;
        page.Resources = new[] { new HtmlRuntimeResource(imageUrl, png, "image/png") };
        var options = new HtmlRenderOptions { ViewportWidth = 816D };
        options.AdditionalStylesheets.Add(
            "main{background-image:url('" + imageUrl.AbsoluteUri + "');background-repeat:no-repeat;background-size:12px 12px}");
        IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                RenderRequests = new[] { HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage,
                    HtmlRenderEncoder.Png, options) }
            });

        Assert.DoesNotContain(result.Capture.Resources, resource => resource.Url == imageUrl);
        Assert.Contains(result.RenderResources, resource => resource.Url == imageUrl && resource.Content.SequenceEqual(png));
        Assert.Contains(result.Outputs[0].Render.Document.Pages.SelectMany(renderPage => renderPage.Visuals)
            .OfType<HtmlRenderImage>(), image => image.Source?.EndsWith(":background-image", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(result.Outputs[0].Render.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public async Task RenderOnlyResourceCannotCrossTheRuntimeOriginBoundary() {
        Uri imageUrl = new("https://unapproved.officeimo.test/background.png");
        byte[] png = OfficePngWriter.EncodeRgba(1, 1, new byte[] { 255, 0, 0, 255 });
        HtmlScriptRequest page = LegacyCase().Item1;
        page.Resources = new[] { new HtmlRuntimeResource(imageUrl, png, "image/png") };
        var options = new HtmlRenderOptions { ViewportWidth = 816D };
        options.AdditionalStylesheets.Add("main{background-image:url('" + imageUrl.AbsoluteUri + "')}");
        IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                RenderRequests = new[] { HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage,
                    HtmlRenderEncoder.Png, options) }
            });

        Assert.DoesNotContain(result.RenderResources, resource => resource.Url == imageUrl);
        Assert.DoesNotContain(result.Outputs[0].Render.Document.Pages.SelectMany(renderPage => renderPage.Visuals)
            .OfType<HtmlRenderImage>(), image => image.Source?.EndsWith(":background-image", StringComparison.Ordinal) == true);
    }

    [Fact]
    public async Task DirectResourceIdentityWinsOverAnotherResourcesRedirectAlias() {
        Uri directUrl = new("https://legacy.officeimo.test/b.png");
        Uri redirectedUrl = new("https://legacy.officeimo.test/a.png");
        byte[] directPng = OfficePngWriter.EncodeRgba(1, 1, new byte[] { 255, 0, 0, 255 });
        byte[] redirectedPng = OfficePngWriter.EncodeRgba(1, 1, new byte[] { 0, 0, 255, 255 });
        HtmlScriptRequest page = LegacyCase().Item1;
        page.Resources = new[] {
            new HtmlRuntimeResource(directUrl, directPng, "image/png"),
            new HtmlRuntimeResource(redirectedUrl, redirectedPng, "image/png", finalUrl: directUrl, redirectCount: 1)
        };
        var options = new HtmlRenderOptions { ViewportWidth = 816D };
        options.AdditionalStylesheets.Add("main{background-image:url('" + directUrl.AbsoluteUri +
            "');background-repeat:no-repeat;background-size:12px 12px}");
        IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                RenderRequests = new[] { HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage,
                    HtmlRenderEncoder.Png, options) }
            });

        HtmlRenderImage background = Assert.Single(result.Outputs[0].Render.Document.Pages
            .SelectMany(renderPage => renderPage.Visuals).OfType<HtmlRenderImage>(),
            image => image.Source?.EndsWith(":background-image", StringComparison.Ordinal) == true);
        Assert.Equal(directPng, background.Bytes);
    }

    private static (HtmlScriptRequest Page, HtmlAutomationRequest[] Actions, string Ready, string ExpectedText) CreateCase(string id) => id switch {
        "vanilla" => VanillaCase(),
        "react-build" => ReactBuildCase(),
        "preact" => PreactCase(),
        "legacy" => LegacyCase(),
        _ => throw new ArgumentOutOfRangeException(nameof(id))
    };

    private static (HtmlScriptRequest, HtmlAutomationRequest[], string, string) VanillaCase() {
        string Read(string name) => ReadFixture("StandaloneApplication", name);
        Uri origin = new("https://application.example/index.html");
        Uri review = new(origin, "/review?title=Quarterly&region=South");
        HtmlRuntimeResource Text(string name, string type) => HtmlRuntimeResource.FromText(new Uri(origin, name), Read(name), type);
        var resources = new List<HtmlRuntimeResource> {
            Text("app.js", "text/javascript"), Text("view.js", "text/javascript"),
            Text("app.css", "text/css"), Text("theme.css", "text/css"),
            Text("data.json", "application/json"), Text("review.js", "text/javascript"),
            HtmlRuntimeResource.FromText(review, Read("review.html"), "text/html; charset=utf-8"),
            new(new Uri(origin, "health.txt"), Encoding.UTF8.GetBytes(Read("health.txt")), "text/plain", statusCode: 503)
        };
        return (new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin,
            Html = Read("index.html"), Resources = resources,
            ReadyExpression = "window.applicationReady===true", Timeout = TimeSpan.FromSeconds(30)
        }, new[] {
            Fill("Report title", "Quarterly"), Select("Region", "South"), Click("Prepare review"),
            Wait("h1", "Review report")
        }, "window.reviewReady===true", "Total: 25");
    }

    private static (HtmlScriptRequest, HtmlAutomationRequest[], string, string) ReactBuildCase() {
        string Read(string name) => ReadFixture("ReactBuild", name);
        Uri origin = new("https://built-app.officeimo.test/index.html");
        var resources = new List<HtmlRuntimeResource> {
            HtmlRuntimeResource.FromText(new Uri(origin, "style.css"), Read("style.css"), "text/css"),
            HtmlRuntimeResource.FromText(new Uri(origin, "data.json"), Read("data.json"), "application/json")
        };
        string dist = Path.Combine(AppContext.BaseDirectory, "Fixtures", "ReactBuild", "dist");
        resources.AddRange(Directory.EnumerateFiles(dist, "*.js").Select(path =>
            HtmlRuntimeResource.FromText(new Uri(origin, Path.GetFileName(path)), File.ReadAllText(path), "text/javascript")));
        return (new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin,
            Html = Read("index.html"), Resources = resources,
            ReadyExpression = "document.querySelector('#total')?.textContent==='Total: 42'", Timeout = TimeSpan.FromSeconds(30)
        }, new[] {
            Select("Region", "South"), Wait("#total", "Total: 18"),
            Fill("Report title", "Quarterly"), Click("Add adjustment"),
            Wait("#total", "Total: 21"), Click("Prepare review"), Wait("h1", "Review report")
        }, "document.querySelector('h1')?.textContent==='Review report'", "Total: 21");
    }

    private static (HtmlScriptRequest, HtmlAutomationRequest[], string, string) PreactCase() {
        Uri origin = new("https://application.example/");
        var resources = new[] { "preact.umd.js", "hooks.umd.js", "report.js" }.Select(name =>
            HtmlRuntimeResource.FromText(new Uri(origin, name), ReadFixture("Preact", name), "text/javascript")).ToList();
        resources.Add(HtmlRuntimeResource.FromText(new Uri(origin, "data.json"),
            "[{\"name\":\"North\",\"value\":24},{\"name\":\"South\",\"value\":18}]", "application/json"));
        return (new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin,
            Html = "<!doctype html><div id='app'></div><script src='/preact.umd.js'></script><script src='/hooks.umd.js'></script><script src='/report.js'></script>",
            Resources = resources, Scripts = new[] { "localStorage.setItem('report:name','Monthly');mountReport()" },
            ReadyExpression = "document.querySelector('#total')?.textContent==='Total: 42'", Timeout = TimeSpan.FromSeconds(30)
        }, new[] {
            Select("Region", "South"), Wait("#total", "Total: 18"),
            Fill("Report name", "Quarterly"), Wait("#report-name", "Report: Quarterly")
        }, "document.querySelector('#report-name')?.textContent==='Report: Quarterly'", "Report: Quarterly");
    }

    private static (HtmlScriptRequest, HtmlAutomationRequest[], string, string) LegacyCase() =>
        (new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://legacy.officeimo.test/index.html"),
            Html = ReadFixture("LegacyApplication", "index.html")
        }, new[] { Click("Approve"), Wait("#state", "Approved") },
        "document.querySelector('#state')?.textContent==='Approved'", "Approved");

    private static string ReadFixture(string directory, string name) => File.ReadAllText(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", directory, name));

    private static HtmlAutomationRequest Fill(string name, string value) => new() {
        Query = HtmlLocatorQuery.ByAccessibleName(name), Action = HtmlAutomationAction.Fill, Value = value
    };
    private static HtmlAutomationRequest Select(string name, string value) => new() {
        Query = HtmlLocatorQuery.ByAccessibleName(name), Action = HtmlAutomationAction.SelectOptions, Values = new[] { value }
    };
    private static HtmlAutomationRequest Click(string name) => new() {
        Query = HtmlLocatorQuery.ByAccessibleName(name), Action = HtmlAutomationAction.Click
    };
    private static HtmlAutomationRequest Wait(string selector, string value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Wait,
        WaitState = HtmlLocatorWaitState.Text, Value = value
    };

    private sealed class CountingHost : IHtmlRuntimeHost {
        public HtmlRuntimeProviderDescriptor Descriptor { get; } = new("test", "1", new[] { HtmlRuntimeProfile.ScriptedDocumentV1 },
            new[] { HtmlRuntimeCapabilityIds.OperationTrace }, 1, 1);
        public int ContextsCreated { get; private set; }
        public Task<IHtmlRuntimeContext> CreateContextAsync(HtmlRuntimeContextOptions? options = null, CancellationToken cancellationToken = default) {
            ContextsCreated++;
            throw new InvalidOperationException("A context should not be created for an invalid render request.");
        }
    }
}
