namespace OfficeIMO.Html;

/// <summary>
/// Built-in visual themes for OfficeIMO-generated HTML documents.
/// </summary>
public enum OfficeHtmlDocumentThemeKind {
    /// <summary>Balanced default theme for document-like HTML output.</summary>
    WordLike,

    /// <summary>Compact theme for dense worksheets and technical references.</summary>
    Compact,

    /// <summary>Report theme with stronger headings and table contrast.</summary>
    Report,

    /// <summary>Technical theme for reference-oriented content.</summary>
    Technical
}
