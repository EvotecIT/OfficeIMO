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
    [Fact]
    public void ApplicationRuntimeCreatesTheSupportedProcessHostWithoutProviderWiring() {
        IHtmlRuntimeHost host = HtmlApplicationRuntime.CreateProcessHost(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"));

        Assert.Equal("officeimo.trusted-process", host.Descriptor.Id);
        Assert.Contains(HtmlRuntimeProfile.WebApplicationV1, host.Descriptor.Profiles);
        Assert.True(host.Descriptor.Supports(HtmlRuntimeCapabilityIds.StructuredActions));
    }

    [Fact]
    public void StandardApplicationOutputsCanSelectOneNamedResult() {
        var page = new HtmlScriptRequest {
            ViewportWidth = 1024D,
            ViewportHeight = 640D
        };

        HtmlApplicationDocumentRequest snapshot = new HtmlApplicationDocumentRequest {
            Page = page,
            OutputOptions = new HtmlApplicationOutputOptions {
                Kinds = HtmlApplicationOutputKinds.ScreenToPagePdf
            }
        }.Snapshot();

        HtmlRenderRequest output = Assert.Single(snapshot.RenderRequests);
        Assert.Equal(HtmlRenderIntentProfile.ScreenSnapshotPaged, output.Profile);
        Assert.Equal(HtmlRenderEncoder.Pdf, output.Encoder);
        Assert.Equal(1024D, output.Options.ViewportWidth);
        Assert.Equal(640D, output.Options.ViewportHeight);
        Assert.Equal(HtmlRenderMargins.All(0D), output.Options.Margins);
        Assert.Equal(HtmlRenderUserAgentStyleMode.Browser, output.Options.UserAgentStyles);
        Assert.Equal("serif", output.Options.DefaultFontFamily);
    }

    [Fact]
    public void StandardApplicationOutputsRejectAnEmptySelection() {
        var request = new HtmlApplicationDocumentRequest {
            OutputOptions = new HtmlApplicationOutputOptions { Kinds = HtmlApplicationOutputKinds.None }
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => request.Snapshot());
    }

    [Fact]
    public async Task NamedApplicationResultSkipsARequestWithCustomizedIntentAxes() {
        var page = new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://named-output.officeimo.test/"),
            Html = "<p>Named output</p>"
        };
        var options = new HtmlRenderOptions { ViewportWidth = 320D };
        HtmlRenderRequest customized = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, options)
            .WithCssMedia(HtmlCssMediaContext.Print);
        HtmlRenderRequest standard = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, options);
        IHtmlRuntimeHost host = HtmlApplicationRuntime.CreateProcessHost(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"));

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                RenderRequests = new[] { customized, standard }
            });

        Assert.False(result.Outputs[0].Render.Request.MatchesNamedProfile);
        Assert.Same(result.Outputs[1], result.ScreenPng);
        Assert.Same(result.Outputs[1], result.FindOutput(
            HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png));
    }

    [Fact]
    public void StaticRendererProjectsAuthoredSrcdocIntoAClippedFrameViewport() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("""
            <main>Before <iframe width="220" height="80" srcdoc="<p id='inside'>Static frame body</p>"></iframe> After</main>
            """);

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, new HtmlRenderOptions {
            ViewportWidth = 600D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains("Static frame body", rendered.Text, StringComparison.Ordinal);
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals), visual =>
            visual is HtmlRenderClipGroup { ClipWidth: 220D, ClipHeight: 80D }
            && visual.Source?.EndsWith(":frame-viewport", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void StaticRendererDiagnosesAFrameBeyondTheConfiguredDepth() {
        string nested = "<p>Nested frame</p>";
        nested = $"<iframe srcdoc=\"{System.Net.WebUtility.HtmlEncode(nested)}\"></iframe>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            $"<iframe srcdoc=\"{System.Net.WebUtility.HtmlEncode(nested)}\"></iframe>");

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, new HtmlRenderOptions {
            MaxFrameDepth = 1
        });

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.FrameDepthLimitExceeded
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain("Nested frame", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void SiblingFramesShareTheOperationWideBackgroundTileBudget() {
        string svg = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='1'><rect width='1' height='1' fill='red'/></svg>"));
        string child = $"<div style=\"width:2px;height:2px;background:url(data:image/svg+xml;base64,{svg}) 0 0/1px 1px repeat\"></div>";
        string encoded = System.Net.WebUtility.HtmlEncode(child);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            $"<iframe width='2' height='2' srcdoc=\"{encoded}\"></iframe><iframe width='2' height='2' srcdoc=\"{encoded}\"></iframe>");

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, new HtmlRenderOptions {
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D),
            MaxBackgroundImageTiles = 6
        });

        Assert.Contains(rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded);
    }

    [Fact]
    public async Task CapturedFrameBodyRemainsSeparateAndRendersAcrossAllDocumentIntents() {
        Uri origin = new("https://frames.officeimo.test/index.html");
        Uri frameUrl = new(origin, "/detail.html");
        Uri frameCssUrl = new(origin, "/frame.css");
        var page = new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = origin,
            Html = """
                <style>body{margin:0} iframe{width:260px;height:90px;border:2px solid #24476b}</style>
                <main><h1>Outer report</h1><iframe src="/detail.html"></iframe></main>
                """,
            Resources = new[] {
                HtmlRuntimeResource.FromText(frameUrl, """
                    <link rel="stylesheet" href="/frame.css">
                    <p id="inside">Captured frame body</p>
                    """, "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(frameCssUrl,
                    "body{margin:0;background:#eef6ff}p{color:#17324d;font-weight:bold}",
                    "text/css")
            },
            ReadyExpression = "document.querySelector('iframe')?.contentDocument?.querySelector('#inside')?.textContent==='Captured frame body'"
        };
        IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);
        var rendering = new HtmlToPdfOptions { ViewportWidth = 640D, Margins = HtmlRenderMargins.All(0D) };

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                FinalReadyExpression = page.ReadyExpression,
                RenderRequests = new[] {
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, rendering),
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, rendering),
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, rendering)
                }
            });

        Assert.Null(result.Capture.Document.QuerySelector("#inside"));
        Assert.Equal("Captured frame body", Assert.Single(result.Capture.Frames).Document.QuerySelector("#inside")!.TextContent);
        Assert.Contains(result.RenderResources, resource => resource.Url == frameCssUrl);
        Assert.All(result.Outputs, output => Assert.Contains("Captured frame body", output.Render.Document.Text, StringComparison.Ordinal));
        Assert.All(result.Outputs, output => Assert.DoesNotContain(output.Render.Diagnostics,
            diagnostic => diagnostic.Code is HtmlRenderDiagnosticCodes.ResourceUnavailable
                or HtmlRenderDiagnosticCodes.ExternalStylesheetPending));
        Assert.NotEmpty(result.Outputs[0].Images);
        foreach (HtmlApplicationRenderOutput output in result.Outputs.Skip(1)) {
            PdfReadDocument pdf = PdfReadDocument.Open(output.Pdf!.ToBytes());
            Assert.Contains("Captured frame body", pdf.ExtractText(), StringComparison.Ordinal);
        }
        if (Environment.GetEnvironmentVariable("OFFICEIMO_APPLICATION_EVIDENCE_DIR") is { Length: > 0 } evidenceRoot) {
            string folder = Path.Combine(evidenceRoot, "frame-content");
            Directory.CreateDirectory(folder);
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen.png"), result.Outputs[0].Images[0].Bytes);
            File.WriteAllBytes(Path.Combine(folder, "officeimo-print.pdf"), result.Outputs[1].Pdf!.ToBytes());
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen-to-page.pdf"), result.Outputs[2].Pdf!.ToBytes());
        }
    }

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
        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                Actions = actions,
                FinalReadyExpression = ready,
                OutputOptions = new HtmlApplicationOutputOptions {
                    RenderOptions = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D) }
                }
            });

        Assert.Equal("officeimo.trusted-process", result.Provider.Id);
        Assert.True(result.Capture.Document.IsReadOnly);
        Assert.Equal(actions.Length, result.Actions.Count);
        Assert.All(result.Actions, action => Assert.Equal(HtmlAutomationStatus.Success, action.Status));
        Assert.Contains(result.Trace.Events, item => item.Kind == HtmlRuntimeEventKind.Capture);
        Assert.Equal(3, result.Outputs.Count);
        Assert.All(result.Outputs, output => Assert.Equal(HtmlRenderDocumentState.RuntimeSnapshot, output.Render.Request.DocumentState));
        Assert.All(result.Outputs, output => {
            Assert.Equal(page.ViewportWidth, output.Render.Request.Options.ViewportWidth);
            Assert.Equal(page.ViewportHeight, output.Render.Request.Options.ViewportHeight);
            Assert.Equal(HtmlRenderUserAgentStyleMode.Browser, output.Render.Request.Options.UserAgentStyles);
        });
        Assert.Same(result.Outputs[0], result.ScreenPng);
        Assert.Same(result.Outputs[1], result.PrintPdf);
        Assert.Same(result.Outputs[2], result.ScreenToPagePdf);
        byte[] png = Assert.Single(result.ScreenPng!.Images).Bytes;
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Null(result.ScreenPng.Pdf);
        foreach (HtmlApplicationRenderOutput output in new[] { result.PrintPdf!, result.ScreenToPagePdf! }) {
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
            File.WriteAllBytes(Path.Combine(folder, "officeimo-print.pdf"), result.PrintPdf!.Pdf!.ToBytes());
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen-to-page.pdf"), result.ScreenToPagePdf!.Pdf!.ToBytes());
        }
    }

    [Theory]
    [InlineData("forms")]
    [InlineData("tables")]
    [InlineData("external-graph")]
    [InlineData("navigation")]
    public async Task BroaderApplicationCorpusRetainsFinalStateAcrossStandardOutputs(string caseId) {
        ApplicationCorpusCase application = CreateApplicationCorpusCase(caseId);
        IHtmlRuntimeHost host = HtmlApplicationRuntime.CreateProcessHost(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"));

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = application.Page,
                Actions = application.Actions,
                FinalReadyExpression = application.FinalReadyExpression,
                OutputOptions = new HtmlApplicationOutputOptions {
                    RenderOptions = new HtmlToPdfOptions {
                        ViewportWidth = 816D,
                        ViewportHeight = 720D,
                        Margins = HtmlRenderMargins.All(0D)
                    }
                }
            });

        Assert.Equal(application.Actions.Count, result.Actions.Count);
        Assert.All(result.Actions, action => Assert.Equal(HtmlAutomationStatus.Success, action.Status));
        Assert.Equal(3, result.Outputs.Count);
        Assert.All(result.Outputs, output => {
            Assert.Equal(HtmlRenderDocumentState.RuntimeSnapshot, output.Render.Request.DocumentState);
            Assert.Contains(application.ExpectedText, output.Render.Document.Text, StringComparison.Ordinal);
        });
        byte[] png = Assert.Single(result.ScreenPng!.Images).Bytes;
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
        Assert.NotNull(image);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_APPLICATION_EVIDENCE_DIR") is { Length: > 0 } evidenceRoot) {
            string folder = Path.Combine(evidenceRoot, "html-application-corpus", caseId);
            Directory.CreateDirectory(folder);
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen.png"), png);
            File.WriteAllBytes(Path.Combine(folder, "officeimo-print.pdf"), result.PrintPdf!.Pdf!.ToBytes());
            File.WriteAllBytes(Path.Combine(folder, "officeimo-screen-to-page.pdf"), result.ScreenToPagePdf!.Pdf!.ToBytes());
        }
        foreach (HtmlApplicationRenderOutput output in new[] { result.PrintPdf!, result.ScreenToPagePdf! }) {
            PdfReadDocument pdf = PdfReadDocument.Open(output.Pdf!.ToBytes());
            Assert.Contains(application.ExpectedText, pdf.ExtractText(), StringComparison.Ordinal);
            Assert.True(pdf.HasTaggedContent);
        }

        switch (caseId) {
            case "forms":
                Assert.Equal("Quarterly", result.Capture.Document.QuerySelector("#title")!.FormState!.Value);
                Assert.True(result.Capture.Document.QuerySelector("#approved")!.FormState!.IsChecked);
                Assert.True(result.Capture.Document.QuerySelector("option[value=South]")!.FormState!.IsSelected);
                Assert.Equal("Captured notes", result.Capture.Document.QuerySelector("#notes")!.FormState!.Value);
                break;
            case "tables":
                Assert.True(result.PrintPdf!.Render.Document.Pages.Count > 1);
                Assert.Equal(44, result.Capture.Document.QuerySelectorAll("#ledger-body tr").Count);
                break;
            case "external-graph":
                Assert.Contains(result.Capture.Resources, resource => resource.Url.AbsolutePath == "/summary.js");
                Assert.Contains(result.Capture.Resources, resource => resource.Url.AbsolutePath == "/data.json");
                Assert.False(result.Capture.Document.QuerySelector("#summarize")!.HasAttribute("disabled"));
                break;
            case "navigation":
                Assert.Equal("https://navigation.officeimo.test/report/approved", result.Capture.DocumentUrl.AbsoluteUri);
                Assert.Contains(result.Capture.Resources, resource => resource.Url.AbsolutePath == "/report.html");
                break;
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

    [Fact]
    public async Task PageDeviceDensityIsInheritedByStaticWorkflowOutputs() {
        Uri origin = new("https://legacy.officeimo.test/index.html");
        Uri twoX = new(origin, "/two-x.svg");
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='blue'/></svg>";
        var page = new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = origin,
            ViewportWidth = 800D,
            ViewportHeight = 600D,
            DevicePixelRatio = 2D,
            Html = "<img src='/fallback.svg' srcset='/two-x.svg 2x' width='24' height='24' alt='fixture'>",
            Resources = new[] { HtmlRuntimeResource.FromText(twoX, svg, "image/svg+xml") }
        };
        IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
            AngleSharpDomServices.Instance);

        HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
            new HtmlApplicationDocumentRequest {
                Page = page,
                RenderRequests = new[] {
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png,
                        new HtmlRenderOptions { ViewportWidth = 800D, ViewportHeight = 600D })
                }
            });

        Assert.Contains(result.Capture.Resources, resource => resource.Url == twoX);
        Assert.Equal(192D, result.Outputs[0].Render.Request.Options.MediaFeatures.ResolutionDpi);
        Assert.DoesNotContain(result.Outputs[0].Render.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
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

    private static ApplicationCorpusCase CreateApplicationCorpusCase(string caseId) {
        string directory = caseId switch {
            "forms" => "Forms",
            "tables" => "Tables",
            "external-graph" => "ExternalGraph",
            "navigation" => "Navigation",
            _ => throw new ArgumentOutOfRangeException(nameof(caseId))
        };
        Uri origin = new($"https://{caseId}.officeimo.test/index.html");
        string root = Path.Combine(AppContext.BaseDirectory, "Fixtures", "ApplicationCorpus", directory);
        string html = File.ReadAllText(Path.Combine(root, "index.html"));
        var resources = Directory.EnumerateFiles(root)
            .Where(path => !Path.GetFileName(path).Equals("index.html", StringComparison.OrdinalIgnoreCase)
                && !Path.GetFileName(path).Equals("browser-actions.json", StringComparison.OrdinalIgnoreCase))
            .Select(path => HtmlRuntimeResource.FromText(
                new Uri(origin, "/" + Path.GetFileName(path)),
                File.ReadAllText(path),
                ContentType(path)))
            .ToArray();

        return caseId switch {
            "forms" => new(new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin, Html = html, Resources = resources,
                ViewportWidth = 816D, ViewportHeight = 720D, Timeout = TimeSpan.FromSeconds(30)
            }, new[] {
                FillCss("#title", "Quarterly"), SelectCss("#region", "South"), CheckCss("#approved", true),
                FillCss("#notes", "Captured notes"), ClickCss("#prepare"), Wait("#summary", "Quarterly | South | approved | Captured notes")
            }, "window.applicationCorpusReady===true", "Quarterly | South | approved | Captured notes"),
            "tables" => new(new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin, Html = html, Resources = resources,
                ViewportWidth = 816D, ViewportHeight = 720D,
                ReadyExpression = "window.applicationCorpusReady===true", Timeout = TimeSpan.FromSeconds(30)
            }, Array.Empty<HtmlAutomationRequest>(), "window.applicationCorpusReady===true", "Retained ledger entry 42"),
            "external-graph" => new(new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin, Html = html, Resources = resources,
                ViewportWidth = 816D, ViewportHeight = 720D,
                ReadyExpression = "window.applicationCorpusInteractive===true", Timeout = TimeSpan.FromSeconds(30)
            }, new[] { Click("Build summary"), Wait("#summary", "Qualified graph total: 51") },
                "window.applicationCorpusReady===true", "Qualified graph total: 51"),
            "navigation" => new(new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = origin, Html = html, Resources = resources,
                ViewportWidth = 816D, ViewportHeight = 720D, Timeout = TimeSpan.FromSeconds(30)
            }, new[] { Click("Open report"), Wait("h1", "Route report"), Click("Approve route"), Wait("#state", "Approved route snapshot") },
                "window.applicationCorpusReady===true", "Approved route snapshot"),
            _ => throw new ArgumentOutOfRangeException(nameof(caseId))
        };
    }

    private static string ContentType(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".html" => "text/html; charset=utf-8",
        ".css" => "text/css; charset=utf-8",
        ".js" => "text/javascript; charset=utf-8",
        ".json" => "application/json; charset=utf-8",
        _ => "application/octet-stream"
    };

    private static string ReadFixture(string directory, string name) => File.ReadAllText(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", directory, name));

    private static HtmlAutomationRequest Fill(string name, string value) => new() {
        Query = HtmlLocatorQuery.ByAccessibleName(name), Action = HtmlAutomationAction.Fill, Value = value
    };
    private static HtmlAutomationRequest Select(string name, string value) => new() {
        Query = HtmlLocatorQuery.ByAccessibleName(name), Action = HtmlAutomationAction.SelectOptions, Values = new[] { value }
    };
    private static HtmlAutomationRequest FillCss(string selector, string value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Fill, Value = value
    };
    private static HtmlAutomationRequest SelectCss(string selector, string value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.SelectOptions, Values = new[] { value }
    };
    private static HtmlAutomationRequest CheckCss(string selector, bool value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.SetChecked, Checked = value
    };
    private static HtmlAutomationRequest ClickCss(string selector) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Click
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

    private sealed record ApplicationCorpusCase(
        HtmlScriptRequest Page,
        IReadOnlyList<HtmlAutomationRequest> Actions,
        string FinalReadyExpression,
        string ExpectedText);
}
