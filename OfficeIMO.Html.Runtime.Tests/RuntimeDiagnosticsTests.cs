using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeDiagnosticsTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"),
        AngleSharpDomServices.Instance);

    [Fact]
    public async Task TraceRecordsExactDynamicReplayConsumptionIdentity() {
        var url = new Uri("https://diagnostics.officeimo.test/submit");
        var request = new HtmlRuntimeFetchRequest(url, new Uri("https://diagnostics.officeimo.test/"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "text/plain;charset=UTF-8" },
            Encoding.UTF8.GetBytes("payload"));
        var replay = new HtmlRuntimeFetchReplay(request, 1, HtmlRuntimeResource.FromText(url, "accepted", "text/plain"));
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync(new HtmlRuntimeContextOptions {
            Trace = new HtmlRuntimeTraceOptions { IncludeUrls = true }
        });
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://diagnostics.officeimo.test/"),
            Html = "<script>fetch('/submit',{method:'POST',body:'payload'}).then(()=>window.done=true)</script>",
            FetchReplays = new[] { replay }
        });
        await page.WaitForAsync("window.done===true");

        HtmlRuntimeTrace trace = page.GetTrace();

        Assert.Contains(trace.Events, item => item.Operation == "fetch-replay" && item.Status == "consumed" &&
            item.ArtifactId == replay.Identity);
    }

    [Fact]
    public async Task TraceIncludesBoundedProviderEventsAndDeterministicArtifactEvidence() {
        var source = new Uri("https://diagnostics.officeimo.test/private-script.js?token=secret");
        var final = new Uri("https://diagnostics.officeimo.test/assets/private-script.js?token=secret");
        var resource = new HtmlRuntimeResource(source,
            Encoding.UTF8.GetBytes("console.info('secret console value');document.body.dataset.loaded='yes'"),
            "text/javascript", finalUrl: final, redirectCount: 1,
            headers: new Dictionary<string, string> { ["Authorization"] = "Bearer secret" });
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync(new HtmlRuntimeContextOptions {
            Trace = new HtmlRuntimeTraceOptions {
                MaxEvents = 64, IncludeUrls = true, IncludeConsoleMessages = true,
                Redactor = value => value.Replace("secret", "redacted", StringComparison.Ordinal)
            }
        });
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://diagnostics.officeimo.test/"),
            Html = "<!doctype html><body><a id='download' href='/report.pdf' download>Download</a><script src='/private-script.js?token=secret'></script></body>",
            Resources = new[] { resource }
        });
        HtmlAutomationResult download = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("#download"), Action = HtmlAutomationAction.Click, WaitForReady = false
        });
        HtmlScriptCapture first = await page.CaptureAsync();
        HtmlScriptCapture second = await page.CaptureAsync();

        Assert.Equal(HtmlAutomationStatus.Unsupported, download.Status);
        Assert.Equal(first.ArtifactManifest.Id, second.ArtifactManifest.Id);
        Assert.Equal(2, first.ArtifactManifest.Entries.Count);
        using JsonDocument serializedCapture = JsonDocument.Parse(HtmlRuntimeJson.Serialize(first));
        string documentHtml = serializedCapture.RootElement.GetProperty("documentHtml").GetString()!;
        string serializedDocumentHash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(documentHtml))).ToLowerInvariant();
        Assert.Equal(serializedDocumentHash, first.ArtifactManifest.Entries.Single(item => item.Name == "document.html").Sha256);
        HtmlRuntimeTrace trace = page.GetTrace();
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Resource && item.StatusCode == 200 && item.RedirectCount == 1);
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Redirect && item.RedirectCount == 1);
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Policy && item.Decision == "supplied");
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Console && item.Detail == "redacted console value");
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Download && item.Decision == "unsupported");
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Artifact && item.ArtifactId == first.ArtifactManifest.Id);
        string json = HtmlRuntimeJson.Serialize(trace);
        Assert.DoesNotContain("secret", json, StringComparison.Ordinal);
        Assert.DoesNotContain("Authorization", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DefaultTraceOmitsUrlsConsoleMessagesAndFailureMessages() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<p>Ready</p>", Scripts = new[] { "console.warn('private-value')" }
        });
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => page.ExecuteAsync("throw new Error('private-failure')"));
        HtmlRuntimeTrace trace = page.GetTrace();
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Console && item.Detail == null);
        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Failure && item.Detail == null);
        Assert.All(trace.Events, item => Assert.Null(item.Url));
        Assert.DoesNotContain("private-value", HtmlRuntimeJson.Serialize(trace), StringComparison.Ordinal);
        Assert.DoesNotContain("private-failure", HtmlRuntimeJson.Serialize(trace), StringComparison.Ordinal);
    }

    [Fact]
    public async Task RedactorCanOmitOptionalTraceValuesWithoutFallingBackToSensitiveInput() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync(new HtmlRuntimeContextOptions {
            Trace = new HtmlRuntimeTraceOptions {
                IncludeUrls = true,
                IncludeConsoleMessages = true,
                Redactor = _ => null!
            }
        });
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri("https://diagnostics.officeimo.test/?token=secret"),
            Html = "<p>Ready</p>",
            Scripts = new[] { "console.info('secret console')" }
        });

        HtmlRuntimeTrace trace = page.GetTrace();

        Assert.Contains(trace.Events, item => item.Kind == HtmlRuntimeEventKind.Console && item.Detail == null);
        Assert.DoesNotContain(trace.Events, item => item.Url != null);
        Assert.DoesNotContain("secret", HtmlRuntimeJson.Serialize(trace), StringComparison.Ordinal);
    }
}
