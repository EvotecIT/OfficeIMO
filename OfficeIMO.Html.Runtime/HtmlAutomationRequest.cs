namespace OfficeIMO.Html.Runtime;

/// <summary>A live-document automation operation.</summary>
public enum HtmlAutomationAction {
    /// <summary>Read one element's current state.</summary>
    Inspect = 0,
    /// <summary>Count matches without a strict single-element requirement.</summary>
    Count = 1,
    /// <summary>Activate an element through DOM click behavior.</summary>
    Click = 2,
    /// <summary>Move the selected primary pointer over an element.</summary>
    Hover = 3,
    /// <summary>Dispatch one selected keyboard key and its qualified default behavior.</summary>
    Press = 4,
    /// <summary>Replace an editable text control's value and dispatch input.</summary>
    Fill = 5,
    /// <summary>Activate a checkbox or radio to obtain the requested checked state.</summary>
    SetChecked = 6,
    /// <summary>Select options by exact, unambiguous values.</summary>
    SelectOptions = 7,
    /// <summary>Focus an element.</summary>
    Focus = 8,
    /// <summary>Blur the element if it owns focus.</summary>
    Blur = 9,
    /// <summary>Scroll the layout viewport by the minimum amount needed to expose the element's bounding box.</summary>
    ScrollIntoView = 10,
    /// <summary>Wait for a locator state without executing caller-supplied JavaScript.</summary>
    Wait = 11,
    /// <summary>Set the selection range of an editable text control.</summary>
    SetSelection = 12
}

/// <summary>Modifier keys carried by a WebApplicationV1 keyboard action.</summary>
[Flags]
public enum HtmlKeyboardModifiers {
    /// <summary>No modifier key.</summary>
    None = 0,
    /// <summary>The Alt key.</summary>
    Alt = 1,
    /// <summary>The Control key.</summary>
    Control = 2,
    /// <summary>The Meta or Command key.</summary>
    Meta = 4,
    /// <summary>The Shift key.</summary>
    Shift = 8
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
    /// <summary>One element has a nonempty box and is visible by the qualified CSS layout model.</summary>
    Visible,
    /// <summary>No element matches, or the single match has no visible box.</summary>
    Hidden,
    /// <summary>One visible element intersects the current layout viewport.</summary>
    InViewport,
    /// <summary>The current control value equals Value.</summary>
    Value,
    /// <summary>Normalized element text equals Value.</summary>
    Text,
    /// <summary>Current checkbox or radio checkedness equals Checked.</summary>
    Checked
}

/// <summary>Provider-neutral automation request. The session validates and snapshots it before queueing.</summary>
public sealed class HtmlAutomationRequest {
    /// <summary>The query to resolve. Supply either this value or <see cref="Reference"/>.</summary>
    public HtmlLocatorQuery? Query { get; init; }
    /// <summary>A revision-bound target returned by <see cref="IHtmlRuntimePage.ObserveAsync"/>.</summary>
    public HtmlObservedElementReference? Reference { get; init; }
    /// <summary>The requested operation.</summary>
    public HtmlAutomationAction Action { get; init; }
    /// <summary>Text for Fill, the key for Press, or a Value/Text wait.</summary>
    public string? Value { get; init; }
    /// <summary>Modifier keys exposed by a Press action.</summary>
    public HtmlKeyboardModifiers Modifiers { get; init; }
    /// <summary>Inclusive selection start for SetSelection.</summary>
    public int? SelectionStart { get; init; }
    /// <summary>Exclusive selection end for SetSelection.</summary>
    public int? SelectionEnd { get; init; }
    /// <summary>Exact option values for SelectOptions. An empty list clears the selection.</summary>
    public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
    /// <summary>Requested checkedness for SetChecked or a Checked wait.</summary>
    public bool? Checked { get; init; }
    /// <summary>The condition used by Wait.</summary>
    public HtmlLocatorWaitState WaitState { get; init; }
    /// <summary>Retry missing or temporarily unavailable targets until the command deadline. Ambiguous or unsupported requests fail immediately.</summary>
    public bool WaitForReady { get; init; } = true;

    /// <summary>Validates and returns a detached request within a provider's input-character budget.</summary>
    public HtmlAutomationRequest Snapshot(int maximumCharacters) {
        if (maximumCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maximumCharacters));
        if ((Query == null) == (Reference == null))
            throw new ArgumentException("Supply exactly one automation target: Query or Reference.");
        if (Reference != null) {
            if (string.IsNullOrWhiteSpace(Reference.PageId) || Reference.Revision <= 0 || Reference.ElementIndex < 0
                || string.IsNullOrWhiteSpace(Reference.ElementName))
                throw new ArgumentException("The observed element reference is invalid.", nameof(Reference));
        }
        if (!Enum.IsDefined(Action) || !Enum.IsDefined(WaitState)) throw new ArgumentException("Unknown automation operation or condition.");
        const HtmlKeyboardModifiers allModifiers = HtmlKeyboardModifiers.Alt | HtmlKeyboardModifiers.Control
            | HtmlKeyboardModifiers.Meta | HtmlKeyboardModifiers.Shift;
        if ((Modifiers & ~allModifiers) != 0 || Action != HtmlAutomationAction.Press && Modifiers != HtmlKeyboardModifiers.None)
            throw new ArgumentException("Keyboard modifiers are valid only for Press actions.");
        if (Action != HtmlAutomationAction.SetSelection && (SelectionStart is not null || SelectionEnd is not null))
            throw new ArgumentException("Selection offsets are valid only for SetSelection actions.");
        if ((Action is HtmlAutomationAction.Fill or HtmlAutomationAction.Press || Action == HtmlAutomationAction.Wait && WaitState is HtmlLocatorWaitState.Value or HtmlLocatorWaitState.Text) && Value == null)
            throw new ArgumentException("The operation requires Value.");
        if (Action == HtmlAutomationAction.SetSelection && (SelectionStart is null || SelectionEnd is null || SelectionStart < 0 || SelectionEnd < 0))
            throw new ArgumentException("SetSelection requires nonnegative start and end offsets.");
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
        characters += Reference?.PageId.Length ?? 0;
        characters += Reference?.ElementName.Length ?? 0;
        characters += Reference?.ElementId.Length ?? 0;
        if (characters > maximumCharacters) throw new ArgumentException("The automation request exceeds MaxInputCharacters.");
        return new() { Query = Query, Reference = Reference, Action = Action, Value = Value, Values = Array.AsReadOnly(values), Checked = Checked,
            Modifiers = Modifiers, SelectionStart = SelectionStart, SelectionEnd = SelectionEnd,
            WaitState = WaitState, WaitForReady = WaitForReady };
    }
}
