namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in visual profiles shared by Markdown exporters.
/// </summary>
public enum MarkdownVisualThemeKind {
    /// <summary>No opinionated document styling beyond each renderer's defaults.</summary>
    Plain,

    /// <summary>Neutral document styling that resembles a clean Word document.</summary>
    WordLike,

    /// <summary>Polished technical-document styling for guides, READMEs, and specifications.</summary>
    TechnicalDocument,

    /// <summary>GitHub-inspired Markdown styling.</summary>
    GitHubLike,

    /// <summary>Compact styling for dense notes and command output.</summary>
    Compact,

    /// <summary>Report-oriented styling with stronger tables and section hierarchy.</summary>
    Report
}
