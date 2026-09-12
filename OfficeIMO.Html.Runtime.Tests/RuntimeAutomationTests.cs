using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeAutomationTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task ReusableLocatorsFillReplacementNodesAndCaptureIndependentValues() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<section><label for='name'>Report name</label><input id='name' value='Default'><button id='done'>Done</button></section>"
        });
        var input = session.Locator(HtmlLocatorQuery.ByAccessibleName("Report name"));
        await input.FillAsync("Quoted 'value'\n\"text\" \\ end");
        var first = await session.CaptureAsync();
        Assert.Equal("Quoted 'value'\"text\" \\ end", (await input.InspectAsync()).Value);
        await session.ExecuteAsync("document.querySelector('#name').outerHTML='<input id=\"name\" value=\"Replacement\">'");
        await input.FillAsync("Final value");
        await session.Locator(HtmlLocatorQuery.ByText("Done")).ClickAsync();
        var last = await session.CaptureAsync();
        await session.DisposeAsync();
        Assert.NotEqual(first.Document.QuerySelector("#name")!.FormState!.Value, last.Document.QuerySelector("#name")!.FormState!.Value);
        Assert.Equal("Final value", last.Document.QuerySelector("#name")!.FormState!.Value);
        Assert.Contains("Final value", PdfReadDocument.Open(HtmlConversionDocument.FromDocument(last.Document).ToPdfBytes()).ExtractText());
    }

    [Fact]
    public async Task FocusEventsActiveElementAndSelectorsShareSessionState() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<section id='scope'><input id='a'><input id='b'></section>",
            Scripts = new[] { "window.events=[];for(const name of ['focus','focusin','blur','focusout','beforeinput','input','change'])document.addEventListener(name,e=>events.push([e.type,e.target.id,document.activeElement&&document.activeElement.id]),true)" }
        });
        await session.Locator("#a").FillAsync("Value");
        Assert.True((await session.EvaluateAsync("document.activeElement===document.querySelector('#a') && document.querySelector('#a').matches(':focus') && document.querySelector('#scope').matches(':focus-within')")).GetBoolean());
        await session.ExecuteAsync("document.querySelector('#b').focus()");
        Assert.True((await session.Locator("#b").InspectAsync()).IsFocused);
        Assert.False((await session.Locator("#a").InspectAsync()).IsFocused);
        Assert.Equal("focus,focusin,beforeinput,input,change,blur,focusout,focus,focusin", (await session.EvaluateAsync("events.map(e=>e[0]).join(',')")).GetString());
        await session.Locator("#b").BlurAsync();
        Assert.True((await session.EvaluateAsync("document.activeElement===document.body && document.querySelectorAll(':focus').length===0")).GetBoolean());
    }

    [Fact]
    public async Task StrictFailuresAreRecoverableAndScopesDeduplicateInDocumentOrder() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<section class='scope'><div class='scope'><button>A</button></div><button>B</button></section>" });
        var query = session.Locator(".scope").Locator("button");
        Assert.Equal(2, await query.CountAsync());
        var failure = await Assert.ThrowsAsync<HtmlAutomationException>(() => query.ClickAsync());
        Assert.Equal(HtmlAutomationStatus.Ambiguous, failure.Result.Status);
        Assert.Equal("B", (await query.Nth(1).InspectAsync()).Text);
        var invalid = await session.AutomateAsync(new() { Query = HtmlLocatorQuery.Css("["), Action = HtmlAutomationAction.Count });
        Assert.Equal(HtmlAutomationStatus.InvalidLocator, invalid.Status);
        Assert.Equal(2, await query.CountAsync());
    }

    [Fact]
    public async Task WaitsResolveFreshTargetsAndDisabledControlsWithoutCallerScripts() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<input disabled id='value'>",
            Scripts = new[] { "setTimeout(()=>{const input=document.querySelector('#value');input.disabled=false;input.value='Ready'},120)" }
        });
        var input = session.Locator("#value");
        await input.WaitForAsync(HtmlLocatorWaitState.Enabled);
        await input.WaitForValueAsync("Ready");
        await input.FillAsync("Changed");
        await input.WaitForValueAsync("Changed");
        await session.ExecuteAsync("setTimeout(()=>document.querySelector('#value').remove(),40)");
        await input.WaitForAsync(HtmlLocatorWaitState.Detached);
        Assert.Equal(0, await input.CountAsync());
    }

    [Fact]
    public async Task CheckingControlsUsesActivationRollbackGroupsAndIdempotence() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<input type='checkbox' id='check'><input type='radio' name='choice' id='a' checked><input type='radio' name='choice' id='b'>",
            Scripts = new[] { "window.events=[];for(const name of ['click','input','change'])document.addEventListener(name,e=>events.push(e.target.id+':'+name));window.block=e=>e.preventDefault()" }
        });
        await session.Locator("#check").SetCheckedAsync(true);
        await session.Locator("#check").SetCheckedAsync(true);
        Assert.Equal("check:click,check:input,check:change", (await session.EvaluateAsync("events.join(',')")).GetString());
        await session.Locator("#b").SetCheckedAsync(true);
        Assert.False((await session.Locator("#a").InspectAsync()).IsChecked);
        Assert.True((await session.Locator("#b").InspectAsync()).IsChecked);
        await session.ExecuteAsync("document.querySelector('#check').addEventListener('click',block)");
        var cancelled = await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("#check").SetCheckedAsync(false));
        Assert.Equal(HtmlAutomationStatus.Rejected, cancelled.Result.Status);
        Assert.True((await session.Locator("#check").InspectAsync()).IsChecked);
    }

    [Fact]
    public async Task OptionSelectionIsOrdinalValidatedBeforeMutationAndCanBeCleared() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<select id='choice' multiple><option value='a'>Alpha</option><option value='A'>Upper</option><option value='b' disabled>Disabled</option></select>"
        });
        var select = session.Locator("#choice");
        await select.SelectOptionsAsync(new[] { "A" });
        Assert.Equal(new[] { "A" }, (await select.InspectAsync()).SelectedValues);
        var rejected = await session.AutomateAsync(new() { Query = select.Query, Action = HtmlAutomationAction.SelectOptions, Values = new[] { "a", "b" }, WaitForReady = false });
        Assert.Equal(HtmlAutomationStatus.NotReady, rejected.Status);
        Assert.Equal(new[] { "A" }, (await select.InspectAsync()).SelectedValues);
        await select.SelectOptionsAsync(Array.Empty<string>());
        Assert.Empty((await select.InspectAsync()).SelectedValues);
    }

    [Fact]
    public async Task PageCancellationAndFocusRedirectionDoNotModifyARejectedFill() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<input id='a' value='Original'><input id='b'>",
            Scripts = new[] { "window.prevent=e=>e.preventDefault();document.querySelector('#a').addEventListener('beforeinput',prevent)" }
        });
        var first = await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("#a").FillAsync("Rejected"));
        Assert.Equal(HtmlAutomationStatus.Rejected, first.Result.Status);
        Assert.Equal("Original", (await session.Locator("#a").InspectAsync()).Value);
        await session.Locator("#a").BlurAsync();
        await session.ExecuteAsync("document.querySelector('#a').removeEventListener('beforeinput',prevent);document.querySelector('#a').addEventListener('focus',()=>document.querySelector('#b').focus())");
        var second = await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("#a").FillAsync("Rejected again"));
        Assert.Equal(HtmlAutomationStatus.Rejected, second.Result.Status);
        Assert.True((await session.Locator("#b").InspectAsync()).IsFocused);
        Assert.Equal("Original", (await session.Locator("#a").InspectAsync()).Value);
    }
}
