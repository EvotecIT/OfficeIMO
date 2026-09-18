using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFetchReplayTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static readonly Uri Origin = new("https://app.example/");
    private static async Task<JsonElement> RunAsync(string script, HtmlScriptRequest? request = null) {
        request ??= new HtmlScriptRequest { DocumentUrl = Origin };
        request.Html = "<p id='status'></p><script>(async()=>{" + script + "})().then(value=>{window.result=value;window.done=true;});</script>";
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.WaitForAsync("window.done===true");
        return await session.EvaluateAsync("window.result");
    }

    [Fact]
    public void DynamicRequestIdentitySnapshotsHeadersBodyAndFetchOptions() {
        byte[] body = Encoding.UTF8.GetBytes("payload");
        var headers = new Dictionary<string, string> { ["X-Variant"] = "blue" };
        var request = new HtmlRuntimeFetchRequest(new Uri(Origin, "submit"), "post", headers, body, credentials: "omit");
        string identity = request.Identity;
        body[0] = (byte)'X';
        headers["X-Variant"] = "changed";

        Assert.Equal("POST", request.Method);
        Assert.Equal("payload", Encoding.UTF8.GetString(request.Body!));
        Assert.Equal("blue", request.Headers["X-Variant"]);
        Assert.Equal(identity, request.Identity);
        Assert.Equal(identity, new HtmlRuntimeFetchRequest(new Uri(Origin, "submit"), "POST",
            new Dictionary<string, string> { ["x-variant"] = "blue" }, Encoding.UTF8.GetBytes("payload"), credentials: "omit").Identity);
        var noBody = new HtmlRuntimeFetchRequest(new Uri(Origin, "empty"), "POST");
        var emptyBody = new HtmlRuntimeFetchRequest(new Uri(Origin, "empty"), "POST", body: Array.Empty<byte>());
        Assert.False(noBody.HasBody);
        Assert.True(emptyBody.HasBody);
        Assert.NotEqual(noBody.Identity, emptyBody.Identity);
        var headeredNoBody = new HtmlRuntimeFetchRequest(new Uri(Origin, "collision"), "POST",
            new Dictionary<string, string> { ["x"] = "y" });
        var headerlessBody = new HtmlRuntimeFetchRequest(new Uri(Origin, "collision"), "POST", body:
            new byte[] { 0, 0, 0, (byte)'x', 1, 0, 0, 0, (byte)'y', 0 });
        Assert.NotEqual(headeredNoBody.Identity, headerlessBody.Identity);
    }

    [Fact]
    public async Task OfflineDiscoveryCarriesExactHeaderedPostAfterScriptHandlesRejection() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = "<script>fetch('/submit',{method:'POST',headers:{'Content-Type':'application/json','X-Variant':'blue'},body:'{\"value\":42}'}).catch(()=>window.handled=true)</script>",
            FailOnFetchReplayDiscovery = true
        }));

        Assert.True(failure.MissingFetchRequests.Count == 1,
            $"Expected one dynamic discovery, got {failure.MissingFetchRequests.Count}: {failure.Message}");
        HtmlRuntimeFetchDiscovery discovery = failure.MissingFetchRequests[0];
        Assert.Empty(failure.MissingResourceUrls);
        Assert.Equal(1, discovery.Occurrence);
        Assert.Equal("POST", discovery.Request.Method);
        Assert.Equal(new Uri(Origin, "submit"), discovery.Request.Url);
        Assert.Equal("blue", discovery.Request.Headers["X-Variant"]);
        Assert.Equal("application/json", discovery.Request.Headers["Content-Type"]);
        Assert.Equal("{\"value\":42}", Encoding.UTF8.GetString(discovery.Request.Body!));
    }

    [Fact]
    public async Task OfflineDiscoveryChargesRequestCountAndRetainedBodyBudgets() {
        var bodyFailure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = "<script>Promise.all([fetch('/one',{method:'POST',body:'ab'}).catch(()=>{}),fetch('/two',{method:'POST',body:'cd'}).catch(()=>{})]).then(()=>window.done=true)</script>",
            ReadyExpression = "window.done===true",
            FailOnFetchReplayDiscovery = true,
            ResourcePolicy = new() { MaxRequestBytes = 2, MaxTotalRequestBytes = 2 }
        }));
        Assert.Single(bodyFailure.MissingFetchRequests);
        Assert.Equal(2, bodyFailure.MissingFetchRequests[0].Request.BodyLength);

        var countFailure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = "<script>Promise.all([fetch('/one',{headers:{'X-Test':'1'}}).catch(()=>{}),fetch('/two',{headers:{'X-Test':'2'}}).catch(()=>{})]).then(()=>window.done=true)</script>",
            ReadyExpression = "window.done===true",
            FailOnFetchReplayDiscovery = true,
            ResourcePolicy = new() { MaxRequests = 1 }
        }));
        Assert.Single(countFailure.MissingFetchRequests);
    }

    [Fact]
    public async Task OfflineDiscoveryStaysWithinTheWorkerResponseCharacterBudget() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = """
                <script>
                async function discover() {
                  await fetch('/one',{method:'POST',body:'x'.repeat(1024)}).catch(()=>{});
                  await fetch('/two',{method:'POST',body:'y'.repeat(1024)}).catch(()=>{});
                  window.done=true;
                }
                discover();
                </script>
                """,
            ReadyExpression = "window.done===true",
            FailOnFetchReplayDiscovery = true,
            MaxOutputCharacters = 4096,
            ResourcePolicy = new() { MaxRequestBytes = 2048, MaxTotalRequestBytes = 4096 }
        }));

        Assert.Contains("response character budget", failure.Message, StringComparison.Ordinal);
        HtmlRuntimeFetchDiscovery retained = Assert.Single(failure.MissingFetchRequests);
        Assert.Equal(new Uri(Origin, "one"), retained.Request.Url);
        Assert.Equal(1024, retained.Request.BodyLength);
    }

    [Fact]
    public async Task ExactDynamicReplaysPreserveRepeatedRequestOccurrences() {
        var request = new HtmlRuntimeFetchRequest(new Uri(Origin, "submit"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "text/plain;charset=UTF-8", ["X-Variant"] = "blue" },
            Encoding.UTF8.GetBytes("payload"));
        var first = new HtmlRuntimeFetchReplay(request, 1,
            HtmlRuntimeResource.FromText(request.Url, "first", "text/plain"));
        var second = new HtmlRuntimeFetchReplay(request, 2,
            HtmlRuntimeResource.FromText(request.Url, "second", "text/plain"));
        var result = await RunAsync("""
            const options={method:'POST',headers:{'Content-Type':'text/plain;charset=UTF-8','X-Variant':'blue'},body:'payload'};
            const first=await (await fetch('/submit',options)).text();
            const second=await (await fetch('/submit',options)).text();
            return {first,second};
            """, new HtmlScriptRequest { DocumentUrl = Origin, FetchReplays = new[] { first, second } });

        Assert.Equal("first", result.GetProperty("first").GetString());
        Assert.Equal("second", result.GetProperty("second").GetString());
    }

    [Fact]
    public async Task HeaderedXmlHttpRequestConsumesAnExactDynamicReplay() {
        var request = new HtmlRuntimeFetchRequest(new Uri(Origin, "api/header-varying.json"), headers:
            new Dictionary<string, string> { ["X-Variant"] = "private" });
        HtmlScriptCapture capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = """
                <p id="result">Loading</p>
                <script>
                  const request = new XMLHttpRequest();
                  request.open('GET','/api/header-varying.json');
                  request.setRequestHeader('X-Variant','private');
                  request.onload = () => document.querySelector('#result').textContent = `Ready ${request.responseText}`;
                  request.onerror = () => document.querySelector('#result').textContent = 'Failed';
                  request.onloadend = () => document.body.dataset.settled = 'yes';
                  request.send();
                </script>
                """,
            ReadyExpression = "document.body.dataset.settled === 'yes'",
            FetchReplays = new[] { new HtmlRuntimeFetchReplay(request, 1,
                HtmlRuntimeResource.FromText(request.Url, "42", "text/plain")) }
        });

        Assert.Contains("Ready 42", capture.Document.Body!.TextContent, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingLaterOccurrenceReportsTheConsumedReplayTranscript() {
        var request = new HtmlRuntimeFetchRequest(new Uri(Origin, "submit"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "text/plain;charset=UTF-8" },
            Encoding.UTF8.GetBytes("payload"));
        var first = new HtmlRuntimeFetchReplay(request, 1,
            HtmlRuntimeResource.FromText(request.Url, "first", "text/plain"));

        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = "<script>const options={method:'POST',headers:{'Content-Type':'text/plain;charset=UTF-8'},body:'payload'};fetch('/submit',options).then(()=>fetch('/submit',options)).catch(()=>window.handled=true)</script>",
            FetchReplays = new[] { first },
            FailOnFetchReplayDiscovery = true
        }));

        Assert.Equal(new[] { first.Identity }, failure.ConsumedFetchReplayIdentities);
        Assert.Equal(2, Assert.Single(failure.MissingFetchRequests).Occurrence);
    }

    [Fact]
    public async Task ExactReplayStillEnforcesSameOriginMode() {
        var request = new HtmlRuntimeFetchRequest(new Uri("https://cdn.example/data"), mode: "same-origin");
        var replay = new HtmlRuntimeFetchReplay(request, 1,
            HtmlRuntimeResource.FromText(request.Url, "private", "text/plain"));
        var result = await RunAsync("""
            try { await fetch('https://cdn.example/data',{mode:'same-origin'}); }
            catch(error) { return String(error); }
            return 'unexpected';
            """, new HtmlScriptRequest { DocumentUrl = Origin, FetchReplays = new[] { replay },
                ResourcePolicy = new() { AllowedOrigins = new[] { new Uri("https://cdn.example/") } } });

        Assert.Contains("same-origin mode", result.GetString());
    }

    [Fact]
    public async Task CrossOriginExactReplayRejectsRequestsRequiringPreflight() {
        var safeRequest = new HtmlRuntimeFetchRequest(new Uri("https://cdn.example/safe"), mode: "cors");
        var unsafeRequest = new HtmlRuntimeFetchRequest(new Uri("https://cdn.example/delete"), "DELETE", mode: "cors");
        var corsHeaders = new Dictionary<string, string> { ["Access-Control-Allow-Origin"] = Origin.GetLeftPart(UriPartial.Authority) };
        var result = await RunAsync("""
            const safe=await (await fetch('https://cdn.example/safe',{mode:'cors'})).text();
            let unsafe; try { await fetch('https://cdn.example/delete',{method:'DELETE',mode:'cors'}); }
            catch(error) { unsafe=String(error); }
            return {safe,unsafe};
            """, new HtmlScriptRequest {
                DocumentUrl = Origin,
                FetchReplays = new[] {
                    new HtmlRuntimeFetchReplay(safeRequest, 1, new HtmlRuntimeResource(safeRequest.Url,
                        Encoding.UTF8.GetBytes("safe"), "text/plain", headers: corsHeaders)),
                    new HtmlRuntimeFetchReplay(unsafeRequest, 1, new HtmlRuntimeResource(unsafeRequest.Url,
                        Array.Empty<byte>(), "text/plain", headers: corsHeaders))
                },
                ResourcePolicy = new() { AllowedOrigins = new[] { new Uri("https://cdn.example/") } }
            });

        Assert.Equal("safe", result.GetProperty("safe").GetString());
        Assert.Contains("preflight", result.GetProperty("unsafe").GetString(), StringComparison.OrdinalIgnoreCase);
    }
}
