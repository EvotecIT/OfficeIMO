using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeBeforeUnloadTests {
    private static readonly Uri Start = new("https://app.example/start");
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData("'unsaved'", true)]
    [InlineData("''", true)]
    [InlineData("false", true)]
    [InlineData("null", false)]
    [InlineData("undefined", false)]
    public async Task HandlerReturnValueControlsHeadlessNavigation(string value, bool canceled) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<body><h1>Start</h1></body>",
            Scripts = ["window.onbeforeunload=() => " + value],
            Resources = [HtmlRuntimeResource.FromText(new Uri(Start, "/next"), "<h1>Next</h1>", "text/html")]
        });
        await session.NavigateAsync(new Uri(Start, "/next"));
        var capture = await session.CaptureAsync();
        Assert.Equal(canceled ? Start : new Uri(Start, "/next"), capture.DocumentUrl);
        if (canceled) {
            await session.ExecuteAsync("window.onbeforeunload=null");
            await session.NavigateAsync(new Uri(Start, "/next"));
            Assert.Equal(new Uri(Start, "/next"), (await session.CaptureAsync()).DocumentUrl);
        }
    }

    [Theory]
    [InlineData("addEventListener('beforeunload', event => { event.returnValue='saved?'; })", true)]
    [InlineData("addEventListener('beforeunload', () => 'ignored')", false)]
    [InlineData("addEventListener('beforeunload', event => { event.returnValue=''; })", false)]
    [InlineData("addEventListener('beforeunload', event => event.preventDefault())", true)]
    [InlineData("document.body.onbeforeunload=()=>''", true)]
    [InlineData("document.body.setAttribute('onbeforeunload', \"return 'inline'\")", true)]
    public async Task ListenerAndBodyHandlersUseTheSameNavigationDecision(string script, bool canceled) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<body><h1>Start</h1></body>", Scripts = [script],
            Resources = [HtmlRuntimeResource.FromText(new Uri(Start, "/next"), "<h1>Next</h1>", "text/html")]
        });
        await session.NavigateAsync(new Uri(Start, "/next"));
        Assert.Equal(canceled ? Start : new Uri(Start, "/next"), (await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task HostAndScriptReplacementUseTypedTrustedEventsWhileSameDocumentRoutesDoNotPrompt() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<body><h1>Start</h1></body>",
            Resources = [HtmlRuntimeResource.FromText(new Uri(Start, "/next"), "<h1>Next</h1>", "text/html")]
        });
        await session.ExecuteAsync("window.calls=0;window.valid=true;onbeforeunload=e=>{calls++;valid=valid && e instanceof BeforeUnloadEvent && e instanceof Event && e.isTrusted && e.cancelable && !e.bubbles && e.returnValue==='';return ''}");
        await session.ExecuteAsync("location.hash='anchor';history.pushState({},'', '#route');history.back()");
        Assert.Equal(0, (await session.EvaluateAsync("calls")).GetInt32());
        await session.ExecuteAsync("location.assign('/next')");
        await session.ReloadAsync();
        await session.ExecuteAsync("location.reload()");
        Assert.Equal(3, (await session.EvaluateAsync("calls")).GetInt32());
        Assert.True((await session.EvaluateAsync("valid && location.pathname==='/start'")).GetBoolean());
        await session.ExecuteAsync("onbeforeunload=null;location.assign('/next')");
        await session.ExecuteAsync("window.calls=0;onbeforeunload=()=>{calls++;return 'stay'};history.back()");
        Assert.Equal(new Uri(Start, "/next"), (await session.CaptureAsync()).DocumentUrl);
        Assert.Equal(1, (await session.EvaluateAsync("calls")).GetInt32());
        await session.ExecuteAsync("onbeforeunload=null;history.back()");
        Assert.Equal("/start", (await session.CaptureAsync()).DocumentUrl!.AbsolutePath);
    }

    [Fact]
    public async Task OrdinarySyntheticEventsDoNotUseTheBeforeUnloadReturnContract() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, Html = "<body></body>"
        });
        Assert.True((await session.EvaluateAsync("(()=>{onbeforeunload=()=>false; const e=new Event('beforeunload',{cancelable:true}); return dispatchEvent(e) && !e.defaultPrevented;})()")).GetBoolean());
    }


    [Fact]
    public async Task HandlerReturnConversionFailuresRemainFatalEvenIfDispatchIsCaught() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, Html = "<body></body>"
        });
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.ExecuteAsync(
            "onbeforeunload=()=>({toString(){throw new Error('conversion failed')}}); try {dispatchEvent(new Event('beforeunload'));}catch(e){};onbeforeunload=null"));
        Assert.Contains("conversion failed", error.Message);
    }


    [Fact]
    public async Task DomParserRemainsSynchronousAndInertInsideAnActiveWorkerScript() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, Html = "<body><h1>Active</h1></body>"
        });
        await session.ExecuteAsync("const parsed=new DOMParser().parseFromString('<script>window.unexpected=true</scr'+'ipt><p>Parsed</p>', 'text/html'); document.querySelector('h1').textContent=parsed.querySelector('p').textContent; window.parsedUrl=parsed.URL");
        Assert.True((await session.EvaluateAsync("document.querySelector('h1').textContent==='Parsed' && typeof unexpected==='undefined' && parsedUrl===location.href")).GetBoolean());
    }

}
