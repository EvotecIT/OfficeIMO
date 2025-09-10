namespace OfficeIMO.Markdown;

/// <summary>
/// Whether to render an embeddable HTML fragment or a standalone HTML document.
/// </summary>
public enum HtmlKind {
    /// <summary>Embeddable fragment (no html/head/body).</summary>
    Fragment,
    /// <summary>Standalone HTML5 document.</summary>
    Document
}

