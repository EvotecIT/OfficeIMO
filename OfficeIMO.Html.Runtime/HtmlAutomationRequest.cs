namespace OfficeIMO.Html.Runtime;

/// <summary>A live-document automation operation.</summary>
public enum HtmlAutomationAction {
    /// <summary>Read one element's current state.</summary>
    Inspect,
    /// <summary>Count matches without a strict single-element requirement.</summary>
    Count,
    /// <summary>Activate an element through DOM click behavior.</summary>
    Click,
    /// <summary>Replace an editable text control's value and dispatch input.</summary>
    Fill,
    /// <summary>Activate a checkbox or radio to obtain the requested checked state.</summary>
    SetChecked,
    /// <summary>Select options by exact, unambiguous values.</summary>
    SelectOptions,
    /// <summary>Focus an element.</summary>
    Focus,
    /// <summary>Blur the element if it owns focus.</summary>
    Blur,
    /// <summary>Wait for a locator state without executing caller-supplied JavaScript.</summary>
    Wait
}

/// <summary>A condition evaluated repeatedly on the document event loop.</summary>
public enum HtmlLocatorWaitState {
    /// <summary>Exactly one element exists.</summary>
    Attached,
    /// <summary>No elements match.</summary>
    Detached,
    /// <summary>One element exists and is enabled.</summary>
    Enabled,
    /// <summary>One element exists and is disabled.</summary>
    Disabled,
    /// <summary>One editable text control exists.</summary>
    Editable,
    /// <summary>One element owns focus.</summary>
    Focused,
    /// <summary>The current control value equals Value.</summary>
    Value,
    /// <summary>Normalized element text equals Value.</summary>
    Text,
    /// <summary>Current checkbox or radio checkedness equals Checked.</summary>
    Checked
}

/// <summary>Provider-neutral automation request. The session validates and snapshots it before queueing.</summary>
public sealed class HtmlAutomationRequest {
    /// <summary>The query to resolve.</summary>
    public HtmlLocatorQuery Query { get; init; } = null!;
    /// <summary>The requested operation.</summary>
    public HtmlAutomationAction Action { get; init; }
    /// <summary>Text for Fill or a Value/Text wait.</summary>
    public string? Value { get; init; }
    /// <summary>Exact option values for SelectOptions. An empty list clears the selection.</summary>
    public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
    /// <summary>Requested checkedness for SetChecked or a Checked wait.</summary>
    public bool? Checked { get; init; }
    /// <summary>The condition used by Wait.</summary>
    public HtmlLocatorWaitState WaitState { get; init; }
    /// <summary>Retry missing or temporarily unavailable targets until the command deadline. Ambiguous or unsupported requests fail immediately.</summary>
    public bool WaitForReady { get; init; } = true;

    internal HtmlAutomationRequest Snapshot(int maximumCharacters) {
        ArgumentNullException.ThrowIfNull(Query);
        if (!Enum.IsDefined(Action) || !Enum.IsDefined(WaitState)) throw new ArgumentException("Unknown automation operation or condition.");
        if ((Action == HtmlAutomationAction.Fill || Action == HtmlAutomationAction.Wait && WaitState is HtmlLocatorWaitState.Value or HtmlLocatorWaitState.Text) && Value == null)
            throw new ArgumentException("The operation requires Value.");
        if ((Action == HtmlAutomationAction.SetChecked || Action == HtmlAutomationAction.Wait && WaitState == HtmlLocatorWaitState.Checked) && Checked == null)
            throw new ArgumentException("The operation requires Checked.");
        ArgumentNullException.ThrowIfNull(Values);
        var values = Values.ToArray();
        long characters = Value?.Length ?? 0;
        foreach (string item in values) {
            if (item == null) throw new ArgumentException("An option value cannot be null.", nameof(Values));
            characters += item.Length + 1L;
        }
        for (var query = Query; query != null; query = query.Scope) characters += query.Value.Length;
        if (characters > maximumCharacters) throw new ArgumentException("The automation request exceeds MaxInputCharacters.");
        return new() { Query = Query, Action = Action, Value = Value, Values = Array.AsReadOnly(values), Checked = Checked, WaitState = WaitState, WaitForReady = WaitForReady };
    }
}
