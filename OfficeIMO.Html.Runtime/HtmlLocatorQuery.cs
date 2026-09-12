namespace OfficeIMO.Html.Runtime;

/// <summary>The syntax used to find elements in a live document.</summary>
public enum HtmlLocatorKind {
    /// <summary>A CSS selector evaluated by the session's syntax provider.</summary>
    Css,
    /// <summary>Normalized descendant text, selecting the smallest matching elements.</summary>
    Text,
    /// <summary>The accessible name resolved by OfficeIMO's HTML and ARIA naming rules.</summary>
    AccessibleName
}

/// <summary>An immutable query, resolved afresh for each operation. It never retains a live node.</summary>
public sealed class HtmlLocatorQuery {
    /// <summary>Creates a query, optionally scoped to matching ancestors and an explicit zero-based result index.</summary>
    public HtmlLocatorQuery(HtmlLocatorKind kind, string value, bool exact = true, HtmlLocatorQuery? scope = null, int? index = null) {
        if (!Enum.IsDefined(kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        ArgumentException.ThrowIfNullOrWhiteSpace(value);
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        int depth = 1;
        for (var parent = scope; parent != null; parent = parent.Scope)
            if (++depth > 16) throw new ArgumentException("Locator scope depth cannot exceed 16.", nameof(scope));
        Kind = kind; Value = value; Exact = exact; Scope = scope; Index = index;
    }
    /// <summary>Query syntax.</summary>
    public HtmlLocatorKind Kind { get; }
    /// <summary>Selector, normalized text, or accessible name.</summary>
    public string Value { get; }
    /// <summary>Whether text and names must match exactly; false uses an ordinal substring match.</summary>
    public bool Exact { get; }
    /// <summary>Optional ancestor scope. Only descendants of matching scope elements are searched.</summary>
    public HtmlLocatorQuery? Scope { get; }
    /// <summary>Optional zero-based index applied after matching and document-order deduplication.</summary>
    public int? Index { get; }
    /// <summary>Creates a CSS locator.</summary>
    public static HtmlLocatorQuery Css(string selector) => new(HtmlLocatorKind.Css, selector);
    /// <summary>Creates a locator for normalized descendant text.</summary>
    public static HtmlLocatorQuery ByText(string text, bool exact = true) => new(HtmlLocatorKind.Text, text, exact);
    /// <summary>Creates a locator using OfficeIMO's bounded accessible-name resolution.</summary>
    public static HtmlLocatorQuery ByAccessibleName(string name, bool exact = true) => new(HtmlLocatorKind.AccessibleName, name, exact);
    /// <summary>Scopes an unscoped query to descendants of another query.</summary>
    public HtmlLocatorQuery Within(HtmlLocatorQuery scope) {
        ArgumentNullException.ThrowIfNull(scope);
        if (Scope != null) throw new InvalidOperationException("This query already has a scope.");
        return new(Kind, Value, Exact, scope, Index);
    }
    /// <summary>Selects an explicit zero-based match without weakening strict single-element actions.</summary>
    public HtmlLocatorQuery Nth(int index) => new(Kind, Value, Exact, Scope, index);
}
