using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeBaseUrlTests {
    private static readonly Uri Start = new("https://base.example/start/page");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData("data:text/plain,ignored")]
    [InlineData("javascript:ignored")]
    [InlineData("https://[")]
    public async Task InvalidBaseFreezesFallbackAcrossHistoryAndCapture(string href) {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = Start, Html = "<base><a href='item'>Item</a>" });
        await session.ExecuteAsync("document.querySelector('base').href=" + System.Text.Json.JsonSerializer.Serialize(href) + ";history.pushState(null,'','/moved/page')");
        Assert.Equal(Start.AbsoluteUri, (await session.EvaluateAsync("document.baseURI")).GetString());
        Assert.Equal(Start, (await session.CaptureAsync()).BaseUri);
        Assert.Equal("https://base.example/start/item", (await session.EvaluateAsync("document.querySelector('a').href")).GetString());
        await session.ExecuteAsync("const base=document.querySelector('base');base.setAttribute('href',base.getAttribute('href'))");
        Assert.Equal("https://base.example/moved/page", (await session.EvaluateAsync("document.baseURI")).GetString());
    }

    [Fact]
    public async Task RelativeScriptAndFetchUseBaseFrozenAtMutationTime() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Start, Html = "<base href='./initial/'><p id='result'>Waiting</p>", Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Start, "/middle/assets/app.js"), "document.querySelector('#result').textContent='Script loaded';window.loaded=true", "text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(Start, "/middle/assets/data.txt"), "Expected data", "text/plain")
            }
        });
        await session.ExecuteAsync("""
            history.pushState(null,'','/middle/page');
            document.querySelector('base').href='./assets/';
            history.pushState(null,'','/final/page');
            const script=document.createElement('script');script.src='app.js';document.body.appendChild(script);
            fetch('data.txt').then(r=>r.text()).then(text=>window.data=text);
            """);
        await session.WaitForAsync("window.loaded===true && window.data==='Expected data'");
        Assert.Equal("https://base.example/middle/assets/", (await session.EvaluateAsync("document.baseURI")).GetString());
        Assert.Equal("Script loaded", (await session.CaptureAsync()).Document.QuerySelector("#result")!.TextContent);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task HostClickAndScriptNavigationUseMutatedBase(bool hostClick) {
        var destination = new Uri(Start, "/middle/assets/next");
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<base href='./initial/'><a id='next' href='next'>Continue</a>",
            Resources = new[] { HtmlRuntimeResource.FromText(destination, "<h1>Correct destination</h1>", "text/html") }
        });
        await session.ExecuteAsync("history.pushState(null,'','/middle/page');document.querySelector('base').href='./assets/';history.pushState(null,'','/final/page')");
        if (hostClick) await session.Locator("#next").ClickAsync();
        else await session.ExecuteAsync("location.assign('next')");
        Assert.Equal(destination, (await session.CaptureAsync()).DocumentUrl);
        Assert.Equal("Correct destination", (await session.CaptureAsync()).Document.QuerySelector("h1")!.TextContent);
    }

    [Fact]
    public async Task SecondaryDocumentBaseCannotRedirectTheActiveDocument() {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = Start, Html = "<base href='/active/'><a href='item'>Item</a>" });
        await session.ExecuteAsync("const other=new DOMParser().parseFromString('<base href=\"https://other.example/assets/\"><a href=\"item\">Other</a>','text/html');window.otherHref=other.querySelector('a').href");
        Assert.Equal("https://other.example/assets/item", (await session.EvaluateAsync("otherHref")).GetString());
        Assert.Equal("https://base.example/active/item", (await session.EvaluateAsync("document.querySelector('a').href")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InertTemplateBasesDoNotRedirectRelativeResources(bool activeBase) {
        var expectedBase = activeBase ? new Uri(Start, "assets/") : Start;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Start, Html = "<template><base href='/inert/'></template>" + (activeBase ? "<base href='./assets/'>" : "") + "<p>Ready</p>",
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(expectedBase, "data.txt"), "Expected data", "text/plain") }
        });
        if (activeBase) await session.ExecuteAsync("history.pushState(null,'','/moved/page')");
        await session.ExecuteAsync("document.querySelector('template').innerHTML='<base href=\"/also-inert/\">';fetch('data.txt').then(r=>r.text()).then(text=>window.data=text)");
        await session.WaitForAsync("window.data==='Expected data'");
        Assert.Equal(expectedBase, (await session.CaptureAsync()).BaseUri);
    }
}
