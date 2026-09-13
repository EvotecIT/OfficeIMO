namespace OfficeIMO.Html.Runtime;

/// <summary>The outcome of an automation operation. Expected failures leave the session usable.</summary>
public enum HtmlAutomationStatus {
    /// <summary>The operation completed.</summary>
    Success,
    /// <summary>No element matched.</summary>
    NotFound,
    /// <summary>A strict operation matched more than one element.</summary>
    Ambiguous,
    /// <summary>The syntax provider rejected the locator.</summary>
    InvalidLocator,
    /// <summary>The target is temporarily unavailable or its wait condition is false.</summary>
    NotReady,
    /// <summary>The control or default action is outside the supported interaction contract.</summary>
    Unsupported,
    /// <summary>The value or option selection is invalid.</summary>
    InvalidValue,
    /// <summary>Page event handlers cancelled or redirected the requested change.</summary>
    Rejected
}

/// <summary>An immutable element rectangle in CSS pixels relative to the current layout viewport.</summary>
public sealed class HtmlRuntimeRect {
    /// <summary>Horizontal viewport coordinate.</summary>
    public double X { get; init; }
    /// <summary>Vertical viewport coordinate.</summary>
    public double Y { get; init; }
    /// <summary>Rectangle width.</summary>
    public double Width { get; init; }
    /// <summary>Rectangle height.</summary>
    public double Height { get; init; }
}

/// <summary>An independent snapshot of one live element's inspected properties.</summary>
public sealed class HtmlRuntimeElementState {
    /// <summary>Element local name.</summary>
    public string ElementName { get; init; } = string.Empty;
    /// <summary>Element id.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>OfficeIMO's bounded accessible name.</summary>
    public string AccessibleName { get; init; } = string.Empty;
    /// <summary>Normalized descendant text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Current input, textarea or select value; null for other elements.</summary>
    public string? Value { get; init; }
    /// <summary>Current selected option values in document order.</summary>
    public IReadOnlyList<string> SelectedValues { get; init; } = Array.Empty<string>();
    /// <summary>Current checkbox/radio state; null for other elements.</summary>
    public bool? IsChecked { get; init; }
    /// <summary>Whether a checkbox is indeterminate.</summary>
    public bool IsIndeterminate { get; init; }
    /// <summary>Whether the element is disabled by form semantics.</summary>
    public bool IsDisabled { get; init; }
    /// <summary>Whether a text control declares applicable readonly state.</summary>
    public bool IsReadOnly { get; init; }
    /// <summary>Whether this element can accept the supported Fill action.</summary>
    public bool IsEditable { get; init; }
    /// <summary>Whether hidden markup or an inert ancestor prevents interaction. This is not a layout-visibility result.</summary>
    public bool IsHiddenByMarkup { get; init; }
    /// <summary>Whether the element owns session focus.</summary>
    public bool IsFocused { get; init; }
    /// <summary>Whether the inspected element remains attached after an action.</summary>
    public bool IsConnected { get; init; }
    /// <summary>Whether WebApplicationV1 computed a nonempty box that is visible by its qualified CSS model; null when layout is not enabled.</summary>
    public bool? IsVisible { get; init; }
    /// <summary>Whether the visible box intersects the current viewport; null when layout is not enabled or no visible box exists.</summary>
    public bool? IsInViewport { get; init; }
    /// <summary>Whether computed pointer-events permits pointer activation; null when layout is not enabled.</summary>
    public bool? AcceptsPointerEvents { get; init; }
    /// <summary>Current bounding box relative to the viewport, or null when the element has no qualified visible box.</summary>
    public HtmlRuntimeRect? BoundingBox { get; init; }
    /// <summary>Current horizontal document scroll offset in CSS pixels.</summary>
    public double ScrollX { get; init; }
    /// <summary>Current vertical document scroll offset in CSS pixels.</summary>
    public double ScrollY { get; init; }
}

/// <summary>Result of an operation, without provider objects or stale element handles.</summary>
public sealed class HtmlAutomationResult {
    /// <summary>Outcome code.</summary>
    public HtmlAutomationStatus Status { get; init; }
    /// <summary>Number of matches at resolution.</summary>
    public int MatchCount { get; init; }
    /// <summary>Human-readable detail for an unsuccessful operation.</summary>
    public string? Message { get; init; }
    /// <summary>Element state when a single target was inspected or acted on.</summary>
    public HtmlRuntimeElementState? Element { get; init; }
    /// <summary>Throws an action exception for a non-success result without terminating the session.</summary>
    public HtmlAutomationResult EnsureSuccess() {
        if (Status != HtmlAutomationStatus.Success) throw new HtmlAutomationException(this);
        return this;
    }
}

/// <summary>An expected automation failure. The live session remains available.</summary>
public sealed class HtmlAutomationException : Exception {
    /// <summary>Creates an exception from a failed automation result.</summary>
    public HtmlAutomationException(HtmlAutomationResult result) : base(result?.Message ?? "The automation operation failed.") => Result = result ?? throw new ArgumentNullException(nameof(result));
    /// <summary>The failed operation's structured result.</summary>
    public HtmlAutomationResult Result { get; }
}
