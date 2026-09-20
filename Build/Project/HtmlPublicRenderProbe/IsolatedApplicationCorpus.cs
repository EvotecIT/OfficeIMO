using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Rendering;
using OfficeIMO.Pdf;

internal sealed record IsolatedApplicationCorpus(IReadOnlyList<IsolatedApplicationProbeResult> Results) {
    internal static async Task<IsolatedApplicationCorpus> RunAsync(string imageId, string rendererPath,
        string workerPath, string outputDirectory) {
        var results = new List<IsolatedApplicationProbeResult>();
        foreach (ApplicationCase fixture in Cases()) {
            IsolatedApplicationProbeResult result = await RunCaseAsync(
                fixture, imageId, rendererPath, workerPath, outputDirectory);
            results.Add(result);
            if (result.ContainerRemoved != true) break;
        }
        return new IsolatedApplicationCorpus(results.AsReadOnly());
    }

    private static async Task<IsolatedApplicationProbeResult> RunCaseAsync(ApplicationCase fixture,
        string imageId, string rendererPath, string workerPath, string outputDirectory) {
        var stopwatch = Stopwatch.StartNew();
        var connections = new List<string>();
        HtmlIsolatedPublicPageResult? result = null;
        Exception? failure = null;
        bool? containerRemoved = null;
        string? containerName = null;
        await using var server = new ControlledHttpServer(request => Reply(fixture, request));
        var transport = new HtmlPublicAcquisitionTransport(
            (host, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (address, port, token) => {
                connections.Add(address + ":" + port);
                var socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                try {
                    await socket.ConnectAsync(IPAddress.Loopback, server.Port, token);
                    return new NetworkStream(socket, ownsSocket: true);
                } catch { socket.Dispose(); throw; }
            });
        try {
            result = await HtmlIsolatedPublicPageWorkflow.RunWithTransportAsync(
                new HtmlIsolatedPublicPageExecutionOptions {
                    ImageId = imageId,
                    PodmanCommand = "podman",
                    PublishedRendererAssemblyPath = rendererPath,
                    PublishedWorkerAssemblyPath = workerPath,
                    OperationTimeout = TimeSpan.FromMinutes(5)
                },
                new HtmlIsolatedPublicPageRequest {
                    ScenarioId = "isolated-application-" + fixture.Id,
                    Url = fixture.DocumentUrl,
                    SourceLicense = "OfficeIMO repository fixture (MIT)",
                    Actions = fixture.Actions,
                    FinalReadyExpression = fixture.FinalReadyExpression,
                    Runtime = new HtmlScriptRequest {
                        Profile = HtmlRuntimeProfile.WebApplicationV1,
                        ReadyExpression = fixture.InitialReadyExpression,
                        ViewportWidth = 816D,
                        ViewportHeight = 720D,
                        Timeout = TimeSpan.FromSeconds(10),
                        SessionTimeout = TimeSpan.FromMinutes(2),
                        ResourcePolicy = new HtmlRuntimeResourcePolicy { MaxRequests = 32 }
                    }
                }, transport);
            containerName = result.ContainerName;
            containerRemoved = result.ContainerRemoved;
            if (result.Actions.Count != fixture.Actions.Length || result.Actions.Any(action => action.Status != HtmlAutomationStatus.Success))
                throw new IOException("The isolated application action evidence is incomplete.");
            HtmlIsolatedPageOutput screen = result.Outputs.Single(output => output.Name == "screen.png");
            HtmlIsolatedPageOutput print = result.Outputs.Single(output => output.Name == "print.pdf");
            HtmlIsolatedPageOutput screenToPage = result.Outputs.Single(output => output.Name == "screen-to-page.pdf");
            if (!screen.Content.Span.StartsWith(new byte[] { 137, 80, 78, 71 })
                || !print.Content.Span.StartsWith("%PDF-"u8) || !screenToPage.Content.Span.StartsWith("%PDF-"u8))
                throw new IOException("The isolated application outputs have invalid signatures.");
            if (!PdfReadDocument.Open(print.Content.ToArray()).ExtractText().Contains(fixture.ExpectedText, StringComparison.Ordinal)
                || !PdfReadDocument.Open(screenToPage.Content.ToArray()).ExtractText().Contains(fixture.ExpectedText, StringComparison.Ordinal))
                throw new IOException("The isolated application PDFs do not retain the expected final state.");
            string directory = Path.Combine(outputDirectory, "isolated-applications", fixture.Id);
            Directory.CreateDirectory(directory);
            foreach (HtmlIsolatedPageOutput output in result.Outputs)
                await File.WriteAllBytesAsync(Path.Combine(directory, output.Name), output.Content.ToArray());
        } catch (HtmlIsolatedPublicPageException error) {
            failure = error;
            containerName = error.Evidence.ContainerName;
            containerRemoved = error.Evidence.ContainerRemoved;
        } catch (Exception error) {
            failure = error;
        }
        bool passed = failure == null && result != null && result.ContainerRemoved;
        return new IsolatedApplicationProbeResult(fixture.Id, passed, containerName, containerRemoved,
            stopwatch.ElapsedMilliseconds, result?.ProviderId, result?.CaptureManifest,
            result?.Actions.Count ?? 0, result?.Resources.Count ?? 0, server.Requests.ToArray(), connections.ToArray(),
            result?.FontPackage.Id, result?.FontPackage.ManifestSha256, result?.FontPackage.FilesSha256,
            result?.Outputs.ToDictionary(output => output.Name, output => output.Sha256, StringComparer.Ordinal),
            failure?.GetType().Name, failure?.Message);
    }

    private static ControlledHttpReply Reply(ApplicationCase fixture, ControlledHttpRequest request) {
        if (request.Method != "GET") return ControlledHttpReply.NotFound();
        string name = request.Path == "/" || request.Path == "/index.html"
            ? "index.html" : request.Path.TrimStart('/');
        string path = Path.Combine(fixture.Directory, name.Replace('/', Path.DirectorySeparatorChar));
        string full = Path.GetFullPath(path);
        if (!full.StartsWith(Path.GetFullPath(fixture.Directory) + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase) || !File.Exists(full)
            || Path.GetFileName(full).Equals("browser-actions.json", StringComparison.OrdinalIgnoreCase))
            return ControlledHttpReply.NotFound();
        return new ControlledHttpReply(File.ReadAllBytes(full), ContentType(full), 200);
    }

    private static IReadOnlyList<ApplicationCase> Cases() {
        string root = Path.Combine(AppContext.BaseDirectory, "ApplicationCorpus");
        return new[] {
            new ApplicationCase("forms", Path.Combine(root, "Forms"), new Uri("http://forms.officeimo.test/index.html"),
                "document.readyState === 'complete'", "window.applicationCorpusReady===true",
                new[] {
                    Fill("#title", "Quarterly"), Select("#region", "South"), Check("#approved", true),
                    Fill("#notes", "Captured notes"), ClickCss("#prepare"),
                    Wait("#summary", "Quarterly | South | approved | Captured notes")
                }, "Quarterly | South | approved | Captured notes"),
            new ApplicationCase("tables", Path.Combine(root, "Tables"), new Uri("http://tables.officeimo.test/index.html"),
                "window.applicationCorpusReady===true", "window.applicationCorpusReady===true",
                Array.Empty<HtmlAutomationRequest>(), "Retained ledger entry 42"),
            new ApplicationCase("external-graph", Path.Combine(root, "ExternalGraph"),
                new Uri("http://external-graph.officeimo.test/index.html"),
                "window.applicationCorpusInteractive===true", "window.applicationCorpusReady===true",
                new[] { ClickName("Build summary"), Wait("#summary", "Qualified graph total: 51") },
                "Qualified graph total: 51")
        };
    }

    private static string ContentType(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".html" => "text/html; charset=utf-8",
        ".css" => "text/css; charset=utf-8",
        ".js" => "text/javascript; charset=utf-8",
        ".json" => "application/json; charset=utf-8",
        _ => "application/octet-stream"
    };
    private static HtmlAutomationRequest Fill(string selector, string value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Fill, Value = value
    };
    private static HtmlAutomationRequest Select(string selector, string value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.SelectOptions, Values = new[] { value }
    };
    private static HtmlAutomationRequest Check(string selector, bool value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.SetChecked, Checked = value
    };
    private static HtmlAutomationRequest ClickCss(string selector) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Click
    };
    private static HtmlAutomationRequest ClickName(string name) => new() {
        Query = HtmlLocatorQuery.ByAccessibleName(name), Action = HtmlAutomationAction.Click
    };
    private static HtmlAutomationRequest Wait(string selector, string value) => new() {
        Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Wait,
        WaitState = HtmlLocatorWaitState.Text, Value = value
    };
}

internal sealed record ApplicationCase(string Id, string Directory, Uri DocumentUrl,
    string InitialReadyExpression, string FinalReadyExpression, HtmlAutomationRequest[] Actions, string ExpectedText);

internal sealed record IsolatedApplicationProbeResult(string Name, bool Passed, string? ContainerName,
    bool? ContainerRemoved, long ElapsedMilliseconds, string? ProviderId, string? CaptureManifest,
    int ActionCount, int AcquiredResourceCount, string[] ServerRequests, string[] Connections,
    string? FontPackageId, string? FontManifestSha256, string? FontFilesSha256,
    IReadOnlyDictionary<string, string>? Outputs, string? ErrorKind, string? Error);
