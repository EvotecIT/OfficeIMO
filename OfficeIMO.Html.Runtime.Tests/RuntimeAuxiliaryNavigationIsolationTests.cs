using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeAuxiliaryNavigationIsolationTests {
    private static readonly Uri Start = new("https://popup.example/start");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task CrossOriginPopupRejectsEventAccessIncludingBorrowedMethods() {
        var other = new Uri("https://other.example/next");
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            ResourcePolicy = new() { AllowedOrigins = [new Uri("https://other.example/")] },
            Resources = [HtmlRuntimeResource.FromText(other, "<!doctype html><body>Other", "text/html")],
            Html = "<!doctype html><body>Root"
        });
        await session.ExecuteAsync("window.popup=open('','persistent')");
        await session.NavigateAsync(other);
        var results = await session.EvaluateAsync("""
            (()=>{const popup=open('','persistent'),handler=()=>{};
                return [()=>popup.addEventListener('message',handler),
                    ()=>EventTarget.prototype.addEventListener.call(popup,'message',handler),
                    ()=>popup.removeEventListener('message',handler),
                    ()=>popup.dispatchEvent(new Event('message')),
                    ()=>{popup.onmessage=handler},()=>popup.onmessage,
                    ()=>popup.setTimeout(handler,1),()=>setInterval.call(popup,handler,1)
                ].map(action=>{try{action();return 'allowed'}catch(e){return e.name}})
            })()
            """);
        Assert.All(results.EnumerateArray(), result => Assert.Equal("SecurityError", result.GetString()));
    }

    [Fact]
    public async Task SavedOpenerEventMethodsTargetTheReplacementRoot() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, Html = "<!doctype html><body>Root"
        });
        await session.ExecuteAsync("""
            window.popup=open('','persistent');const script=popup.document.createElement('script');
            script.textContent="const saved=opener;addEventListener('message',()=>{saved.addEventListener('proof',()=>saved.document.body.setAttribute('data-event','yes'));saved.document.body.setAttribute('data-installed','yes')})";
            popup.document.body.append(script);
            """);
        await session.ReloadAsync();
        await session.ExecuteAsync("open('','persistent').postMessage('install','*')");
        await session.WaitForAsync("document.body.hasAttribute('data-installed')");
        await session.ExecuteAsync("dispatchEvent(new Event('proof'))");
        Assert.True((await session.EvaluateAsync("document.body.hasAttribute('data-event')")).GetBoolean());
    }

    [Fact]
    public async Task PendingPopupModuleDoesNotBlockRootReload() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, Html = "<!doctype html><body>Root"
        });
        await session.ExecuteAsync("""
            window.popup=open('','pending');const script=popup.document.createElement('script');script.type='module';
            script.textContent="document.body.setAttribute('data-started','yes');await new Promise(()=>{})";popup.document.body.append(script);
            """);
        await session.WaitForAsync("popup.document.body.hasAttribute('data-started')");
        await session.ReloadAsync();
        Assert.True((await session.EvaluateAsync("open('','pending').document.body.hasAttribute('data-started')")).GetBoolean());
    }
}
