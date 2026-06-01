namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Built-in visual profiles for Markdown to PDF export.
/// </summary>
public enum MarkdownPdfThemeKind {
    /// <summary>No document-level theme; block styles stay intentionally plain.</summary>
    Plain,

    /// <summary>Neutral Word-like document rhythm and typography.</summary>
    WordLike,

    /// <summary>Polished technical-document styling for guides, specs, and README exports.</summary>
    TechnicalDocument,

    /// <summary>GitHub-inspired Markdown document styling.</summary>
    GitHubLike,

    /// <summary>Compact styling for dense notes and command output.</summary>
    Compact,

    /// <summary>Report-oriented styling with stronger tables and section hierarchy.</summary>
    Report
}
