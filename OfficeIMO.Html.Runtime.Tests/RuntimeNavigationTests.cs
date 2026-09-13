using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeNavigationTests {
    private static readonly Uri First = new("https://navigation.example/first");
    private static readonly Uri Second = new(First, "/second");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static string Page(string name, string script = "") => "<!doctype html><meta charset='utf-8'><h1 id='title'>" + name + "</h1><script>" + script + "</script>";
    private static HtmlRuntimeResource Resource(Uri url, string html) => HtmlRuntimeResource.FromText(url, html, "text/html; charset=utf-8");
    private static HtmlScriptRequest Application(string html) => new() { Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = First, Html = html };

    [Fact]
    public async Task HostNavigationRetainsTheSessionAndRebindsExistingLocators() {
        string initial = Page("First", "window.oldGlobal=42;sessionStorage.setItem('name','stored')");
        var request = Application(initial);
        request.Resources = new[] { Resource(Second, Page("Second", "document.querySelector('#title').textContent=sessionStorage.getItem('name')")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        var title = session.Locator("#title");
        var before = await session.CaptureAsync();
        await session.NavigateAsync(Second);
        await title.WaitForTextAsync("stored");
        Assert.True((await session.EvaluateAsync("typeof oldGlobal==='undefined' && history.length===2 && document.readyState==='complete'")).GetBoolean());
        Assert.Equal(First, before.DocumentUrl);
        Assert.Equal("First", before.Document.QuerySelector("#title")!.TextContent);
        Assert.Equal(Second, (await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task CrossDocumentTraversalRestoresGraphsInTheCurrentRealm() {
        string initial = Page("First", "window.popped=0;onpopstate=()=>popped++");
        var request = Application(initial);
        request.Resources = new[] { Resource(First, initial), Resource(Second, Page("Second")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.ExecuteAsync("const key={id:42};const state={key,map:new Map([[key,7]]),date:new Date(123),bytes:new Uint16Array([42]),big:99n};state.self=state;history.replaceState(state,'');location.assign('/second')");
        Assert.Equal(Second, (await session.CaptureAsync()).DocumentUrl);
        await session.ExecuteAsync("history.back()");
        await session.WaitForAsync("location.pathname==='/first' && window.popped===1");
        Assert.True((await session.EvaluateAsync("history.state.self===history.state && history.state.map instanceof Map && history.state.map.get(history.state.key)===7 && history.state.date instanceof Date && history.state.date.getTime()===123 && history.state.bytes instanceof Uint16Array && history.state.bytes[0]===42 && history.state.big===99n && history.length===2")).GetBoolean());
        await session.ExecuteAsync("history.forward()");
        await session.WaitForAsync("location.pathname==='/second'");
    }

    [Fact]
    public async Task TraversalAndReloadRestoreOwnedViewportOffsetsWhenAutomatic() {
        const string tall = "<style>body{margin:0}.space{height:1800px}</style><div class='space'></div>";
        string initial = Page("First", "window.popped=0;onpopstate=()=>popped++") + tall;
        var request = Application(initial);
        request.ViewportWidth = 320;
        request.ViewportHeight = 160;
        request.Resources = new[] { Resource(First, initial), Resource(Second, Page("Second") + tall) };
        await using var session = await Runtime().OpenTrustedAsync(request);

        await session.ExecuteAsync("scrollTo(0,600);location.assign('/second')");
        Assert.Equal(0D, (await session.EvaluateAsync("scrollY")).GetDouble());

        await session.ExecuteAsync("scrollTo(0,350);history.back()");
        await session.WaitForAsync("location.pathname==='/first' && popped===1");
        Assert.Equal(600D, (await session.EvaluateAsync("scrollY")).GetDouble());

        await session.ExecuteAsync("history.forward()");
        await session.WaitForAsync("location.pathname==='/second'");
        Assert.Equal(350D, (await session.EvaluateAsync("scrollY")).GetDouble());

        await session.ExecuteAsync("history.scrollRestoration='manual';scrollTo(0,450);history.back()");
        await session.WaitForAsync("location.pathname==='/first'");
        await session.ExecuteAsync("scrollTo(0,700);history.forward()");
        await session.WaitForAsync("location.pathname==='/second'");
        Assert.Equal(0D, (await session.EvaluateAsync("scrollY")).GetDouble());

        await session.ExecuteAsync("history.scrollRestoration='auto';scrollTo(0,500);location.reload()");
        Assert.Equal(500D, (await session.EvaluateAsync("scrollY")).GetDouble());

        await session.ExecuteAsync("scrollTo(0,650);history.go(0)");
        Assert.Equal(650D, (await session.EvaluateAsync("scrollY")).GetDouble());
    }

    [Fact]
    public async Task ReloadPreservesHistoryStateAndStorageButCreatesFreshGlobals() {
        string initial = Page("First", "window.runs=Number(localStorage.getItem('runs')||0)+1;localStorage.setItem('runs',runs)");
        var request = Application(initial);
        request.Resources = new[] { Resource(First, initial) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.ExecuteAsync("history.replaceState({id:42},'');window.oldGlobal=true;location.reload=()=>{throw new Error('page replacement')}");
        await session.ReloadAsync();
        Assert.True((await session.EvaluateAsync("runs===2 && typeof oldGlobal==='undefined' && history.state.id===42 && history.length===1")).GetBoolean());
        await session.ExecuteAsync("history.go(0)");
        await session.WaitForAsync("window.runs===3");
    }

    [Fact]
    public async Task OfflineInitialSourceCanReloadAndRestoreWithoutBeingDuplicatedAsAResource() {
        string initial = Page("First", "window.runs=Number(sessionStorage.getItem('runs')||0)+1;sessionStorage.setItem('runs',runs)");
        var request = Application(initial);
        request.Resources = new[] { Resource(Second, Page("Second")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.ReloadAsync();
        Assert.Equal(2, (await session.EvaluateAsync("runs")).GetInt32());
        await session.ExecuteAsync("history.replaceState({route:true},'', '/routed')");
        await session.NavigateAsync(Second);
        await session.ExecuteAsync("history.back()");
        await session.WaitForAsync("location.pathname==='/routed'");
        Assert.True((await session.EvaluateAsync("runs===3 && history.state.route===true && history.length===2")).GetBoolean());
    }

    [Fact]
    public async Task NavigationPartitionsStorageAndUsesTheNewDocumentOriginForFetch() {
        var other = new Uri("https://other.example/page");
        string initial = Page("First");
        var request = Application(initial);
        request.ResourcePolicy = new() { AllowedOrigins = new[] { new Uri("https://other.example") } };
        request.Resources = new[] { Resource(First, initial), Resource(other, Page("Other", "window.empty=sessionStorage.getItem('value')===null;sessionStorage.setItem('value','other');fetch('/data',{mode:'same-origin'}).then(r=>r.text()).then(value=>window.fetched=value)")), HtmlRuntimeResource.FromText(new Uri(other, "/data"), "other-data", "text/plain") };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.ExecuteAsync("sessionStorage.setItem('value','first');localStorage.setItem('key','first')");
        await session.NavigateAsync(other);
        await session.WaitForAsync("window.fetched==='other-data'");
        Assert.True((await session.EvaluateAsync("empty && localStorage.getItem('key')===null && sessionStorage.getItem('value')==='other'")).GetBoolean());
        await session.NavigateAsync(First);
        Assert.True((await session.EvaluateAsync("sessionStorage.getItem('value')==='first' && localStorage.getItem('key')==='first'")).GetBoolean());
    }

    [Fact]
    public async Task TimerNavigationCanCompleteWhileAHostWaitIsPending() {
        var request = Application(Page("First"));
        request.Resources = new[] { Resource(Second, Page("Second")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.ExecuteAsync("setTimeout(()=>location.href='/second',50)");
        await session.Locator("#title").WaitForTextAsync("Second");
        Assert.Equal(Second, (await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task NavigationRequestedByTheLoadingPageSupersedesThatPage() {
        var third = new Uri(First, "/third");
        var request = Application(Page("First"));
        request.Resources = new[] {
                Resource(Second, Page("Second", "location.replace('/third')")),
                Resource(third, Page("Third", "window.loaded='third'"))
            };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.NavigateAsync(Second);
        Assert.Equal(third, (await session.CaptureAsync()).DocumentUrl);
        Assert.True((await session.EvaluateAsync("loaded==='third' && history.length===2")).GetBoolean());
    }

    [Fact]
    public async Task SameDocumentHostNavigationDoesNotConsumeDocumentLoadBudget() {
        var request = Application(Page("First"));
        request.MaxNavigations = 1;
        request.Resources = new[] { Resource(First, Page("First reloaded")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.NavigateAsync(new Uri(First, "#details"));
        Assert.Equal(new Uri(First, "#details"), (await session.CaptureAsync()).DocumentUrl);
        await session.ReloadAsync();
        Assert.Equal("First reloaded", (await session.Locator("#title").InspectAsync()).Text);
    }

    [Fact]
    public async Task RedirectedNavigationUsesTheFinalDocumentUrlAndBase() {
        var redirect = new HtmlRuntimeResource(new Uri(First, "/redirect"), Array.Empty<byte>(), "text/html", 302,
            headers: new Dictionary<string,string> { ["Location"] = "/second" });
        var request = Application(Page("First"));
        request.Resources = new[] { redirect, Resource(Second, Page("Second")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.NavigateAsync(redirect.Url);
        var capture = await session.CaptureAsync();
        Assert.Equal(Second, capture.DocumentUrl);
        Assert.Equal(Second, capture.BaseUri);
        Assert.Equal(2, (await session.EvaluateAsync("history.length")).GetInt32());
    }

    [Fact]
    public async Task NoContentResponseLeavesTheCurrentDocumentAndHistoryActive() {
        var request = Application(Page("First"));
        request.Resources = new[] { new HtmlRuntimeResource(Second, Array.Empty<byte>(), "text/html", 204) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.NavigateAsync(Second);
        Assert.Equal(First, (await session.CaptureAsync()).DocumentUrl);
        Assert.True((await session.EvaluateAsync("history.length===1 && document.querySelector('#title').textContent==='First'")).GetBoolean());
    }

    [Fact]
    public async Task CachelessNotModifiedNavigationIsRejected() {
        var request = Application(Page("First"));
        request.Resources = new[] { new HtmlRuntimeResource(Second, Array.Empty<byte>(), "text/html", 304) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.NavigateAsync(Second));
        Assert.Contains("requires an HTTP cache", failure.Message);
    }

    [Fact]
    public async Task OversizedNavigationCannotExecuteBeforeAdmission() {
        int effects = 0;
        await using var server = new RuntimeHttpFixture((path, _) => {
            if (path == "/effect") { Interlocked.Increment(ref effects); return Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")); }
            string script = "<script>fetch('/effect',{method:'POST',body:'ran'})</script>";
            return Task.FromResult(RuntimeHttpFixture.Reply.Text(script + new string('x', 512), "text/html"));
        });
        var request = Application(Page("First"));
        request.MaxInputCharacters = 256;
        request.ResourcePolicy = new() { AllowNetwork = true };
        request.DocumentUrl = server.Origin;
        await using var session = await Runtime().OpenTrustedAsync(request);
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.NavigateAsync(new Uri(server.Origin, "/oversized")));
        await Task.Delay(100);
        Assert.Equal(0, Volatile.Read(ref effects));
    }

    [Fact]
    public async Task NavigationRejectsHtmlOutsideTheUtf8ApplicationProfile() {
        var request = Application(Page("First"));
        request.Resources = new[] { new HtmlRuntimeResource(Second, new byte[] { 0xff }, "text/html") };
        await using var session = await Runtime().OpenTrustedAsync(request);
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.NavigateAsync(Second));
        Assert.Contains("valid UTF-8 HTML", failure.Message);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task NavigatingDoesNotResetResourceOrNavigationBudgets(bool resourceBudget) {
        var request = Application(Page("First"));
        request.MaxNavigations = resourceBudget ? 10 : 1;
        request.ResourcePolicy = new() { MaxRequests = resourceBudget ? 1 : 10 };
        request.Resources = new[] { Resource(Second, Page("Second")) };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.NavigateAsync(Second);
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.ReloadAsync());
        Assert.Contains(resourceBudget ? "Resource request budget" : "navigation count budget", failure.Message);
    }

    [Fact]
    public async Task ScriptedDocumentProfileRejectsDocumentReplacementWithoutTerminatingTheSession() {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First, Html = Page("First") });
        await Assert.ThrowsAsync<NotSupportedException>(() => session.NavigateAsync(Second));
        await Assert.ThrowsAsync<NotSupportedException>(() => session.ReloadAsync());
        await session.ExecuteAsync("window.navigationError='';try{location.assign('/second')}catch(e){navigationError=e.name}");
        Assert.Equal("NotSupportedError", (await session.EvaluateAsync("navigationError")).GetString());
        Assert.Equal("First", (await session.Locator("#title").InspectAsync()).Text);
    }
}
