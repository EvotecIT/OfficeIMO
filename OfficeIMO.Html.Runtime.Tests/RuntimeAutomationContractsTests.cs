using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeAutomationContractsTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData("Total cost")]
    [InlineData("Total\u00a0cost")]
    public async Task NameQueryNormalizesNonbreakingSpacesOnBothOperands(string query) {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<input aria-label='Total&nbsp;cost'>" });
        var result = await session.AutomateAsync(new() { Query = HtmlLocatorQuery.ByAccessibleName(query), WaitForReady = false });
        Assert.Equal(HtmlAutomationStatus.Success, result.Status);
    }

    [Theory]
    [InlineData("focus", "e.target.hidden=true")]
    [InlineData("focusin", "e.target.disabled=true")]
    [InlineData("focus", "e.target.parentNode.setAttribute('inert','')")]
    public async Task DirectFocusRejectsTargetsInvalidatedByHandlers(string eventType, string mutation) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<section><input></section>",
            Scripts = new[] { "document.querySelector('input').addEventListener('" + eventType + "',e=>{" + mutation + "})" }
        });
        var result = await session.AutomateAsync(new() { Query = HtmlLocatorQuery.Css("input"), Action = HtmlAutomationAction.Focus });
        Assert.Equal(HtmlAutomationStatus.Rejected, result.Status);
        Assert.False((await session.Locator("input").InspectAsync()).IsFocused);
        Assert.True((await session.EvaluateAsync("document.activeElement===document.body && document.querySelectorAll(':focus').length===0 && document.querySelectorAll(':focus-within').length===0")).GetBoolean());
    }

    [Fact]
    public async Task QueuedRequestsSnapshotOptionsAndCancellationDoesNotChangeTheDocument() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<select multiple><option value='a'>A</option><option value='b'>B</option></select><input value='Original'>"
        });
        await session.ExecuteAsync("setTimeout(()=>window.ready=true,300)");
        Task active = session.WaitForAsync("window.ready===true");
        var values = new List<string> { "a" };
        Task<HtmlAutomationResult> queued = session.AutomateAsync(new() { Query = HtmlLocatorQuery.Css("select"), Action = HtmlAutomationAction.SelectOptions, Values = values });
        values[0] = "b";
        using var cancellation = new CancellationTokenSource();
        Task cancelled = session.Locator("input").FillAsync("Wrong", cancellation.Token);
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => cancelled);
        await active;
        (await queued).EnsureSuccess();
        Assert.Equal(new[] { "a" }, (await session.Locator("select").InspectAsync()).SelectedValues);
        Assert.Equal("Original", (await session.Locator("input").InspectAsync()).Value);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ActiveLocatorWaitCancellationOrTimeoutTerminatesSession(bool cancel) {
        await using var session = await Runtime().OpenTrustedAsync(new() { Timeout = TimeSpan.FromSeconds(2) });
        using var cancellation = new CancellationTokenSource();
        Task wait = session.Locator("#missing").WaitForAsync(cancellationToken: cancellation.Token);
        if (cancel) {
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => wait);
        } else await Assert.ThrowsAsync<TimeoutException>(() => wait);
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
    }

    [Fact]
    public async Task RequestValidationPreservesSessionAndListenerFailureTerminatesIt() {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<button>Go</button>", MaxInputCharacters = 256 });
        await Assert.ThrowsAsync<ArgumentException>(() => session.Locator("button").FillAsync(new string('x', 256)));
        await Assert.ThrowsAsync<ArgumentException>(() => session.AutomateAsync(new() { Query = HtmlLocatorQuery.Css("button"), Action = HtmlAutomationAction.SetChecked }));
        Assert.Equal(1, await session.Locator("button").CountAsync());
        await session.ExecuteAsync("document.querySelector('button').onclick=()=>{throw new Error('action failure')}");
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.Locator("button").ClickAsync());
        Assert.Contains("action failure", failure.Message);
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
    }

    [Theory]
    [InlineData("document.querySelector('select').removeAttribute('multiple')")]
    [InlineData("document.querySelector('option').value='changed'")]
    [InlineData("document.querySelector('select').setAttribute('hidden','')")]
    public async Task FocusMutationCannotApplyAStaleSelection(string mutation) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<select multiple><option value='a'>A</option><option value='b'>B</option></select>",
            Scripts = new[] { "document.querySelector('select').addEventListener('focus',()=>{" + mutation + "})" }
        });
        var failure = await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("select").SelectOptionsAsync(new[] { "a", "b" }));
        Assert.Equal(HtmlAutomationStatus.Rejected, failure.Result.Status);
        Assert.Empty((await session.Locator("select").InspectAsync()).SelectedValues);
    }

    [Fact]
    public async Task FocusAndBeforeInputMutationsCannotActivateHiddenOrNonTextControls() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<button>Go</button><input value='Original'>",
            Scripts = new[] { "window.clicks=0;const button=document.querySelector('button');button.onfocus=()=>button.hidden=true;button.onclick=()=>clicks++;document.querySelector('input').addEventListener('beforeinput',e=>e.target.type='checkbox')" }
        });
        Assert.Equal(HtmlAutomationStatus.Rejected, (await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("button").ClickAsync())).Result.Status);
        Assert.Equal(0, (await session.EvaluateAsync("clicks")).GetInt32());
        Assert.Equal(HtmlAutomationStatus.Rejected, (await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("input").FillAsync("Wrong"))).Result.Status);
        Assert.Equal("Original", (await session.Locator("input").InspectAsync()).Value);
    }

    [Fact]
    public async Task AccessibleNamesResolveExplicitImplicitAndAriaLabelsWithoutDuplicateIdMisassociation() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<label for='a'>First</label><label for='a'>Second</label><input id='a'><input id='a' title='Duplicate'>"
                + "<label>Wrapped<input id='b'><input id='c' title='Unlabelled'></label><label for='d'>Ignored</label><input id='d' aria-label='Override'>"
        });
        Assert.Equal("a", (await session.Locator(HtmlLocatorQuery.ByAccessibleName("First Second")).InspectAsync()).Id);
        Assert.Equal("Duplicate", (await session.Locator("input").Nth(1).InspectAsync()).AccessibleName);
        Assert.Equal("b", (await session.Locator(HtmlLocatorQuery.ByAccessibleName("Wrapped")).InspectAsync()).Id);
        Assert.Equal("Unlabelled", (await session.Locator("#c").InspectAsync()).AccessibleName);
        Assert.Equal("Override", (await session.Locator("#d").InspectAsync()).AccessibleName);
    }

    [Fact]
    public async Task AlreadyCheckedMixedControlIsNotToggledAndNavigationDefaultsAreExplicit() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<input type='checkbox' checked><a href='/next'>Next</a>",
            Scripts = new[] { "document.querySelector('input').indeterminate=true;window.clicks=0;document.querySelector('input').onclick=()=>clicks++" }
        });
        await session.Locator("input").SetCheckedAsync(true);
        Assert.True((await session.Locator("input").InspectAsync()).IsIndeterminate);
        Assert.Equal(0, (await session.EvaluateAsync("clicks")).GetInt32());
        Assert.Equal(HtmlAutomationStatus.Unsupported, (await Assert.ThrowsAsync<HtmlAutomationException>(() => session.Locator("a").ClickAsync())).Result.Status);
        await session.ExecuteAsync("document.querySelector('a').onclick=e=>e.preventDefault()");
        await session.Locator("a").ClickAsync();
    }
}
