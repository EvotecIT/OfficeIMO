using System.Net;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFrameBaseUrlTests {
    private static readonly Uri Start = new("https://frames.example/reports/start.html");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task SrcdocResolvesClassicModuleAndFetchAgainstInheritedBase() {
        const string child = """
            <!doctype html><body><a href='next'>Next</a><output id='result'>Waiting</output>
            <script src='classic.js'></script><script type='module'>import {value} from './module.js';document.body.dataset.module=value;</script>
            <script>fetch('data.txt').then(r=>r.text()).then(value=>document.querySelector('#result').textContent=value)</script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<base href='/assets/'><iframe id='child' srcdoc=\"" + WebUtility.HtmlEncode(child) + "\"></iframe>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Start,"/assets/classic.js"), "document.body.dataset.classic='ready';parent.document.body.dataset.childUrl=document.URL", "text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(Start,"/assets/module.js"), "export const value='module ready'", "text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(Start,"/assets/data.txt"), "Frame ready", "text/plain")
            }
        });
        await session.WaitForAsync("document.querySelector('iframe').contentDocument.body.dataset.module==='module ready' && document.querySelector('iframe').contentDocument.querySelector('#result').textContent==='Frame ready'");
        Assert.Equal("about:srcdoc", (await session.EvaluateAsync("document.body.dataset.childUrl")).GetString());
        await session.ExecuteAsync("document.querySelector('base').href='/later/';history.pushState(null,'','/moved/page')");
        Assert.Equal("https://frames.example/assets/next", (await session.EvaluateAsync("document.querySelector('iframe').contentDocument.querySelector('a').href")).GetString());
        var capture = await session.CaptureAsync();
        var frame = Assert.Single(capture.Frames);
        Assert.Equal(new Uri("about:srcdoc"), frame.DocumentUrl);
        Assert.Equal(new Uri(Start,"/assets/"), frame.BaseUri);
        Assert.Equal("ready", frame.Document.Body!.GetAttribute("data-classic"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("src='about:blank?query#fragment'")]
    [InlineData(null)]
    public async Task BlankFrameUsesInheritedBaseWithoutChangingItsIdentity(string? attributes) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Start, Html = "<base href='/assets/'><body>" + (attributes == null ? "" : "<iframe " + attributes + "></iframe>") + "</body>"
        });
        if (attributes == null) await session.ExecuteAsync("document.body.appendChild(document.createElement('iframe'))");
        await session.ExecuteAsync("""
            const child=document.querySelector('iframe').contentDocument;
            child.body.innerHTML="<a href='item'>Item</a>";
            const script=child.createElement('script');
            script.textContent="document.body.dataset.base=document.baseURI;parent.document.body.dataset.child='ran'";
            child.body.appendChild(script);
            document.querySelector('base').href='/later/';
            """);
        await session.WaitForAsync("document.body.dataset.child==='ran'");
        Assert.Equal("https://frames.example/assets/item", (await session.EvaluateAsync("document.querySelector('iframe').contentDocument.querySelector('a').href")).GetString());
        Assert.StartsWith("about:blank", (await session.CaptureAsync()).Frames.Single().DocumentUrl.AbsoluteUri);
    }

    [Theory]
    [InlineData("allow-scripts", false)]
    [InlineData("allow-same-origin", false)]
    [InlineData("allow-scripts allow-same-origin", true)]
    public async Task LocalFrameScriptPolicyUsesOriginNotBaseUrl(string sandbox, bool runs) {
        const string child = "<body><a href='item'>Item</a><script>parent.document.body.dataset.child='ran';localStorage.setItem('child','value');parent.postMessage('ready','/')</script>";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Start,
            Html = "<base href='https://cdn.example/assets/'><body><script>addEventListener('message',e=>document.body.dataset.origin=e.origin)</script><iframe sandbox='" + sandbox + "' srcdoc=\"" + WebUtility.HtmlEncode(child) + "\"></iframe>"
        });
        if (runs) {
            await session.WaitForAsync("document.body.dataset.origin==='https://frames.example'");
            Assert.Equal("value", (await session.EvaluateAsync("localStorage.getItem('child')")).GetString());
        }
        Assert.Equal(runs, (await session.EvaluateAsync("document.body.dataset.child==='ran'")).GetBoolean());
        var capture = await session.CaptureAsync();
        if (sandbox == "allow-scripts") Assert.Empty(capture.Frames);
        else {
            var frame = Assert.Single(capture.Frames);
            Assert.Equal(new Uri("about:srcdoc"), frame.DocumentUrl);
            Assert.Equal(new Uri("https://cdn.example/assets/"), frame.BaseUri);
        }
    }

    [Fact]
    public async Task NestedSrcdocKeepsImmediateCreatorsBaseAfterBothParentsChange() {
        string nested = "<a href='item'>Item</a><script>document.body.dataset.base=document.baseURI</script>";
        string child = "<base href='child/'><iframe srcdoc=\"" + WebUtility.HtmlEncode(nested) + "\"></iframe>";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Start,
            Html = "<base href='/assets/'><iframe srcdoc=\"" + WebUtility.HtmlEncode(child) + "\"></iframe>"
        });
        await session.WaitForAsync("document.querySelector('iframe').contentDocument.querySelector('iframe').contentDocument.body.dataset.base==='https://frames.example/assets/child/'");
        await session.ExecuteAsync("document.querySelector('base').href='/later/';document.querySelector('iframe').contentDocument.querySelector('base').href='other/'");
        var capture = await session.CaptureAsync();
        var childCapture = Assert.Single(capture.Frames);
        var nestedCapture = Assert.Single(childCapture.Frames);
        Assert.Equal(new Uri("about:srcdoc"), nestedCapture.DocumentUrl);
        Assert.Equal(new Uri(Start,"/assets/child/"), nestedCapture.BaseUri);
        Assert.Equal(new Uri(Start,"/assets/other/"), childCapture.BaseUri);
    }

    [Theory]
    [InlineData("allow-scripts")]
    [InlineData("allow-same-origin")]
    public async Task ScriptlessAncestorInitializationCannotBypassSandbox(string sandbox) {
        string child = "<iframe srcdoc=\"" + WebUtility.HtmlEncode("<p>Nested</p><script>top.document.body.dataset.leak='ran'</script>") + "\"></iframe>";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Start,
            Html = "<body><iframe sandbox='" + sandbox + "' srcdoc=\"" + WebUtility.HtmlEncode(child) + "\"></iframe>"
        });
        Assert.False((await session.EvaluateAsync("document.body.dataset.leak==='ran'")).GetBoolean());
    }
}
