namespace OfficeIMO.Html;

/// <summary>
/// Media context used by lightweight CSS analysis in shared HTML conversion workflows.
/// </summary>
public enum HtmlCssMediaContext {
    /// <summary>Screen-oriented analysis, matching browser defaults for HTML conversion previews.</summary>
    Screen,

    /// <summary>Print-oriented analysis for PDF and high-fidelity print conversion profiles.</summary>
    Print
}
