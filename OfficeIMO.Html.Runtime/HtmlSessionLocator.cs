namespace OfficeIMO.Html.Runtime;

/// <summary>A session-bound reusable query. Every call resolves fresh matches and uses the session command deadline.</summary>
public sealed class HtmlSessionLocator {
    private readonly IHtmlRuntimeSession _session;
    internal HtmlSessionLocator(IHtmlRuntimeSession session, HtmlLocatorQuery query) { _session = session; Query = query; }
    /// <summary>The immutable query used by this handle.</summary>
    public HtmlLocatorQuery Query { get; }
    /// <summary>Creates a locator for descendants matching a CSS selector.</summary>
    public HtmlSessionLocator Locator(string selector) => new(_session, HtmlLocatorQuery.Css(selector).Within(Query));
    /// <summary>Selects an explicit zero-based match.</summary>
    public HtmlSessionLocator Nth(int index) => new(_session, Query.Nth(index));
    /// <summary>Counts current matches without waiting for one to appear.</summary>
    public async Task<int> CountAsync(CancellationToken cancellationToken = default) => (await Run(new() { Query = Query, Action = HtmlAutomationAction.Count }, cancellationToken)).MatchCount;
    /// <summary>Reads one current element, waiting for it to appear.</summary>
    public async Task<HtmlRuntimeElementState> InspectAsync(CancellationToken cancellationToken = default) => (await Run(new() { Query = Query }, cancellationToken)).Element!;
    /// <summary>Activates the resolved element through the selected profile's click behavior.</summary>
    public Task ClickAsync(CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Click }, cancellationToken);
    /// <summary>Moves the WebApplicationV1 primary pointer over the resolved element.</summary>
    public Task HoverAsync(CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Hover }, cancellationToken);
    /// <summary>Presses one selected key on the resolved WebApplicationV1 element.</summary>
    public Task PressAsync(string key, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Press, Value = key }, cancellationToken);
    /// <summary>Focuses and replaces a text input or textarea value, dispatching beforeinput/input. A later blur commits change.</summary>
    public Task FillAsync(string value, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Fill, Value = value }, cancellationToken);
    /// <summary>Activates a checkbox/radio to reach the requested state; a page cancellation is reported.</summary>
    public Task SetCheckedAsync(bool value, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.SetChecked, Checked = value }, cancellationToken);
    /// <summary>Selects exact option values. Ambiguous values and unsupported multi-selection are reported.</summary>
    public Task SelectOptionsAsync(IReadOnlyList<string> values, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.SelectOptions, Values = values }, cancellationToken);
    /// <summary>Focuses the resolved element.</summary>
    public Task FocusAsync(CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Focus }, cancellationToken);
    /// <summary>Blurs the element if it owns focus.</summary>
    public Task BlurAsync(CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Blur }, cancellationToken);
    /// <summary>Scrolls the WebApplicationV1 layout viewport just enough to expose the resolved element.</summary>
    public Task ScrollIntoViewAsync(CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.ScrollIntoView }, cancellationToken);
    /// <summary>Waits for an attachment, enabled, editable or focus condition.</summary>
    public Task WaitForAsync(HtmlLocatorWaitState state = HtmlLocatorWaitState.Attached, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Wait, WaitState = state }, cancellationToken);
    /// <summary>Waits for an exact current control value.</summary>
    public Task WaitForValueAsync(string value, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Wait, WaitState = HtmlLocatorWaitState.Value, Value = value }, cancellationToken);
    /// <summary>Waits for exact normalized element text.</summary>
    public Task WaitForTextAsync(string value, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Wait, WaitState = HtmlLocatorWaitState.Text, Value = value }, cancellationToken);
    /// <summary>Waits for current checkedness.</summary>
    public Task WaitForCheckedAsync(bool value, CancellationToken cancellationToken = default) => Run(new() { Query = Query, Action = HtmlAutomationAction.Wait, WaitState = HtmlLocatorWaitState.Checked, Checked = value }, cancellationToken);
    /// <summary>Waits for a nonempty box visible under the WebApplicationV1 CSS layout model.</summary>
    public Task WaitForVisibleAsync(CancellationToken cancellationToken = default) => WaitForAsync(HtmlLocatorWaitState.Visible, cancellationToken);
    /// <summary>Waits until no element matches or the single match has no visible layout box.</summary>
    public Task WaitForHiddenAsync(CancellationToken cancellationToken = default) => WaitForAsync(HtmlLocatorWaitState.Hidden, cancellationToken);
    /// <summary>Waits for the visible element to intersect the current WebApplicationV1 viewport.</summary>
    public Task WaitForInViewportAsync(CancellationToken cancellationToken = default) => WaitForAsync(HtmlLocatorWaitState.InViewport, cancellationToken);
    private async Task<HtmlAutomationResult> Run(HtmlAutomationRequest request, CancellationToken token) => (await _session.AutomateAsync(request, token).ConfigureAwait(false)).EnsureSuccess();
}

/// <summary>Creates provider-neutral locator handles over a runtime session.</summary>
public static class HtmlRuntimeLocatorExtensions {
    /// <summary>Creates a CSS locator.</summary>
    public static HtmlSessionLocator Locator(this IHtmlRuntimeSession session, string selector) => Locator(session, HtmlLocatorQuery.Css(selector));
    /// <summary>Creates a locator for an owned query.</summary>
    public static HtmlSessionLocator Locator(this IHtmlRuntimeSession session, HtmlLocatorQuery query) {
        ArgumentNullException.ThrowIfNull(session); ArgumentNullException.ThrowIfNull(query);
        return new(session, query);
    }
}
