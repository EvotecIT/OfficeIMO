using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFormNavigationTests {
    private static readonly Uri Page = new("https://forms.example/start");
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static HtmlRuntimeResource Resource(string relative, string html) => HtmlRuntimeResource.FromText(
        new Uri(Page, relative), html, "text/html; charset=utf-8");

    [Fact]
    public async Task AutomationResetAndGetSubmissionApplyFormDefaults() {
        const string html = """
            <form action='/result' method='get'>
              <input id='query' name='q' value='default'>
              <button id='reset' type='reset'>Reset</button>
              <button id='submit' type='submit'>Submit</button>
            </form>
            <script>
              window.resets=0;
              const form=document.querySelector('form');
              form.addEventListener('reset',()=>{resets++;form.reset()});
              document.querySelector('form').addEventListener('submit',()=>sessionStorage.setItem('submitted','yes'));
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Page,
            Html = html,
            Resources = new[] { Resource("/result?q=report", "<h1>Result</h1><script>document.body.dataset.submitted=sessionStorage.getItem('submitted')</script>") }
        });

        await session.Locator("#query").FillAsync("changed");
        await session.Locator("#reset").ClickAsync();
        Assert.Equal("default", (await session.Locator("#query").InspectAsync()).Value);
        Assert.Equal(1, (await session.EvaluateAsync("resets")).GetInt32());

        await session.Locator("#query").FillAsync("report");
        await session.Locator("#submit").ClickAsync();
        await session.Locator("h1").WaitForTextAsync("Result");
        Assert.Equal("yes", (await session.EvaluateAsync("document.body.dataset.submitted")).GetString());
        Assert.Equal(new Uri(Page, "/result?q=report"), (await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task SubmitterEventAndOverridesDriveGetSubmission() {
        var preview = new Uri(Page, "/preview?mode=preview");
        const string html = """
            <form action='/normal' method='post'>
              <button id='preview' name='mode' value='preview' formaction='/preview' formmethod='get'>Preview</button>
            </form>
            <script>
              document.querySelector('form').addEventListener('submit',event=>sessionStorage.submitter=event.submitter.id);
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Page,
            Html = html,
            Resources = new[] { HtmlRuntimeResource.FromText(preview, "<h1>Preview</h1>", "text/html; charset=utf-8") }
        });

        await session.Locator("#preview").ClickAsync();

        Assert.Equal(preview, (await session.CaptureAsync()).DocumentUrl);
        Assert.Equal("preview", (await session.EvaluateAsync("sessionStorage.submitter")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EmptyActionUsesTheCurrentDocumentDespiteBaseElement(bool submitterOverride) {
        var result = new Uri(Page, "?q=x");
        string action = submitterOverride ? "action='/wrong'" : "action=''";
        string overrideAttribute = submitterOverride ? "formaction=''" : string.Empty;
        string html = $"<base href='/assets/'><form {action} method='get'><input name='q' value='x'><button id='go' {overrideAttribute}>Go</button></form>";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Page,
            Html = html,
            Resources = new[] { HtmlRuntimeResource.FromText(result, "<h1>Current document</h1>", "text/html; charset=utf-8") }
        });

        await session.Locator("#go").ClickAsync();

        Assert.Equal(result, (await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task ScriptFormMethodsShareResetAndSubmissionOwnership() {
        const string html = """
            <form action='/result' method='get'>
              <input id='query' name='q' value='default'>
              <button id='submit' type='submit'>Submit</button>
            </form>
            <script>
              window.events=[];
              const form=document.querySelector('form');
              form.addEventListener('reset',()=>events.push('reset'));
              form.addEventListener('submit',()=>{events.push('submit');sessionStorage.events=events.join(',')});
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Page,
            Html = html,
            Resources = new[] { Resource("/result?q=script", "<h1>Script result</h1>") }
        });

        await session.ExecuteAsync("(()=>{const query=document.querySelector('#query'),submitter=document.querySelector('#submit'),targetForm=document.querySelector('form');query.value='changed';targetForm.reset();query.value='script';targetForm.requestSubmit(submitter)})()");
        await session.Locator("h1").WaitForTextAsync("Script result");
        Assert.Equal("reset,submit", (await session.EvaluateAsync("sessionStorage.events")).GetString());
    }

    [Fact]
    public async Task CancelledFormDefaultsLeaveTheCurrentDocumentAndValuesAvailable() {
        const string html = """
            <form action='/result' method='get'>
              <input id='query' name='q' value='default'>
              <button id='reset' type='reset'>Reset</button>
              <button id='submit' type='submit'>Submit</button>
            </form>
            <script>
              const form=document.querySelector('form');
              form.addEventListener('reset',event=>event.preventDefault());
              form.addEventListener('submit',event=>event.preventDefault());
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Page,
            Html = html
        });

        await session.Locator("#query").FillAsync("preserved");
        await session.Locator("#reset").ClickAsync();
        await session.Locator("#submit").ClickAsync();
        Assert.Equal("preserved", (await session.Locator("#query").InspectAsync()).Value);
        Assert.Equal(Page, (await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task BeforeUnloadCanCancelAndCommittedNavigationFiresPageHideAndUnload() {
        const string html = """
            <h1>Start</h1>
            <script>
              window.block=true;
              addEventListener('beforeunload',event=>{
                sessionStorage.before=String(+(sessionStorage.before||0)+1);
                if(block)event.preventDefault();
              });
              addEventListener('pagehide',event=>sessionStorage.pagehide=String(event.persisted));
              addEventListener('unload',()=>sessionStorage.unload='yes');
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Page,
            Html = html,
            Resources = new[] { Resource("/next", "<h1>Next</h1>") }
        });

        await session.NavigateAsync(new Uri(Page, "/next"));
        Assert.Equal(Page, (await session.CaptureAsync()).DocumentUrl);
        Assert.Equal("1", (await session.EvaluateAsync("sessionStorage.before")).GetString());

        await session.ExecuteAsync("block=false");
        await session.NavigateAsync(new Uri(Page, "/next"));
        await session.Locator("h1").WaitForTextAsync("Next");
        var lifecycle = await session.EvaluateAsync("({before:sessionStorage.before,pagehide:sessionStorage.pagehide,unload:sessionStorage.unload})");
        Assert.Equal("2", lifecycle.GetProperty("before").GetString());
        Assert.Equal("false", lifecycle.GetProperty("pagehide").GetString());
        Assert.Equal("yes", lifecycle.GetProperty("unload").GetString());
    }
}
