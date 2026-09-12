using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Providers;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeResourceTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static readonly Uri Origin = new("https://app.example/");

    [Fact]
    public async Task SuppliedExternalScriptsAndStylesheetsUseTheDocumentUrlAndRemainAvailableAfterDisposal() {
        var script = HtmlRuntimeResource.FromText(new Uri(Origin, "scripts/app.js"), "document.querySelector('#status').textContent='Loaded';window.ready=true;", "text/javascript");
        var css = HtmlRuntimeResource.FromText(new Uri(Origin, "styles/app.css"), "#status{color:rgb(0, 128, 0)}", "text/css");
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "reports/index.html"),
            Html = "<link rel='stylesheet' href='../styles/app.css'><p id='status'>Pending</p><script src='../scripts/app.js'></script>",
            Resources = new[] { script, css }
        });
        var capture = await session.CaptureAsync("window.ready===true");
        await session.DisposeAsync();
        Assert.Equal("Loaded", capture.Document.QuerySelector("#status")!.TextContent);
        Assert.Equal(new Uri(Origin, "reports/index.html"), capture.DocumentUrl);
        Assert.Equal(2, capture.Resources.Count);
        var loaded = Assert.Single(capture.Resources, resource => resource.Url == script.Url);
        byte[] mutable = loaded.Content;
        mutable[0] = 0;
        Assert.Equal(script.Content, loaded.Content);
        Assert.Equal(css.Content, Assert.Single(capture.Resources, resource => resource.Url == css.Url).Content);
    }

    [Fact]
    public async Task NetworkIsOptInAndAllowedOriginsAreCheckedBeforeRequests() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("window.ready=true")));
        var disabled = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/app.js'></script>"
        }));
        Assert.Contains("network loading is disabled", disabled.Message);
        var forbidden = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin, Html = $"<script src='{new Uri(server.Origin, "app.js")}'></script>", ResourcePolicy = new() { AllowNetwork = true }
        }));
        Assert.Contains("origin is not allowed", forbidden.Message);
        Assert.Empty(server.Requests);
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/app.js'></script>", ResourcePolicy = new() { AllowNetwork = true }, ReadyExpression = "window.ready===true"
        });
        Assert.Single(server.Requests);
        Assert.Equal(new Uri(server.Origin, "app.js"), Assert.Single(capture.Resources).Url);
    }

    [Fact]
    public async Task RedirectsRecheckOriginsAndRetainFinalIdentity() {
        await using var target = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("window.ready=true")));
        await using var source = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("", status: 302, headers: "Location: " + new Uri(target.Origin, "final.js") + "\r\n")));
        var request = new HtmlScriptRequest { DocumentUrl = source.Origin, Html = "<script src='/start.js'></script>", ResourcePolicy = new() { AllowNetwork = true }, ReadyExpression = "window.ready===true" };
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(request));
        Assert.Contains("origin is not allowed", failure.Message);
        Assert.Empty(target.Requests);
        request.ResourcePolicy.AllowedOrigins = new[] { target.Origin };
        var result = Assert.Single((await Runtime().CaptureTrustedAsync(request)).Resources);
        Assert.Equal(new Uri(source.Origin, "start.js"), result.Url);
        Assert.Equal(new Uri(target.Origin, "final.js"), result.FinalUrl);
        Assert.Equal(1, result.RedirectCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ResponseLimitsApplyToDeclaredAndStreamedBytes(bool chunked) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text(new string(' ', 129), chunked: chunked)));
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/large.js'></script>", ResourcePolicy = new() { AllowNetwork = true, MaxResourceBytes = 128 }
        }));
        Assert.Contains("byte budget", failure.Message);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task RequestAndTotalByteBudgetsSpanTheLiveSession(bool requests) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("window.loaded=true;")));
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/first.js'></script>",
            ResourcePolicy = new() { AllowNetwork = true, MaxRequests = requests ? 1 : 10, MaxTotalBytes = requests ? 4096 : 20 }
        });
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(async () => {
            await session.ExecuteAsync("const script=document.createElement('script');script.src='/second.js';document.head.appendChild(script);");
            await session.CaptureAsync("false");
        });
        Assert.Contains(requests ? "request budget" : "byte budget", failure.Message);
    }

    [Fact]
    public async Task CallerCancellationInterruptsAResourceLoad() {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (_, token) => {
            started.TrySetResult();
            await Task.Delay(System.Threading.Timeout.Infinite, token);
            return RuntimeHttpFixture.Reply.Text("");
        });
        using var cancel = new CancellationTokenSource();
        Task capture = Runtime().CaptureTrustedAsync(new HtmlScriptRequest { DocumentUrl = server.Origin, Html = "<script src='/slow.js'></script>", ResourcePolicy = new() { AllowNetwork = true } }, cancel.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancel.Cancel();
        var failure = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => capture);
        Assert.Equal(cancel.Token, failure.CancellationToken);
    }

    [Theory]
    [InlineData("file:///tmp/private.js")]
    [InlineData("data:text/javascript,window.ready=true")]
    public async Task OtherSchemesDoNotBypassTheRequester(string url) {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest { Html = $"<script src='{url}'></script>", ResourcePolicy = new() { AllowNetwork = true } }));
        Assert.Contains("HTTP(S)", failure.Message);
    }

    [Fact]
    public async Task ResourceDeadlineAndHttpFailurePreventSuccessfulCaptures() {
        await using var slow = new RuntimeHttpFixture(async (_, token) => { await Task.Delay(1000, token); return RuntimeHttpFixture.Reply.Text(""); });
        var timeout = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = slow.Origin, Html = "<script src='/slow.js'></script>", ResourcePolicy = new() { AllowNetwork = true, Timeout = TimeSpan.FromMilliseconds(100) }
        }));
        Assert.Contains("resource load exceeded its deadline", timeout.Message);
        await using var missing = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("not found", status: 404)));
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = missing.Origin, Html = "<script src='/missing.js'></script>", ResourcePolicy = new() { AllowNetwork = true }
        }));
        Assert.Contains("HTTP 404", failure.Message);
    }

    [Fact]
    public async Task RedirectLoopsAndUrlCredentialsAreRejected() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("", status: 302, headers: "Location: /again.js\r\n")));
        var loop = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/again.js'></script>", ResourcePolicy = new() { AllowNetwork = true, MaxRedirects = 1 }
        }));
        Assert.Contains("redirect budget", loop.Message);
        Assert.Equal(2, server.Requests.Count);
        var credentialUrl = new UriBuilder(server.Origin) { UserName = "test", Password = "test", Path = "/credentials.js" }.Uri;
        var credentials = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = $"<script src='{credentialUrl}'></script>", ResourcePolicy = new() { AllowNetwork = true }
        }));
        Assert.Contains("without credentials", credentials.Message);
        Assert.Equal(2, server.Requests.Count);
    }

    [Fact]
    public async Task SuppliedResourceIdentityAndBudgetsAreValidatedBeforeOpening() {
        var upper = HtmlRuntimeResource.FromText(new Uri(Origin, "A.js"), "window.upper=true", "text/javascript");
        var lower = HtmlRuntimeResource.FromText(new Uri(Origin, "a.js"), "window.lower=true", "text/javascript");
        var request = new HtmlScriptRequest { DocumentUrl = Origin, Html = "<script src='/A.js'></script><script src='/a.js'></script>", Resources = new[] { upper, lower }, ReadyExpression = "window.upper&&window.lower" };
        Assert.Equal(2, (await Runtime().CaptureTrustedAsync(request)).Resources.Count);
        request.Resources = new[] { upper, upper };
        await Assert.ThrowsAsync<ArgumentException>(() => Runtime().OpenTrustedAsync(request));
        request.Resources = new[] { upper };
        request.ResourcePolicy.MaxResourceBytes = 1;
        await Assert.ThrowsAsync<ArgumentException>(() => Runtime().OpenTrustedAsync(request));
    }
}
