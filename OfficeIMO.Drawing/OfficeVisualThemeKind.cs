namespace OfficeIMO.Drawing;

/// <summary>Built-in visual profiles shared by OfficeIMO document renderers and converters.</summary>
public enum OfficeVisualThemeKind {
    /// <summary>No opinionated document styling beyond each renderer's defaults.</summary>
    Plain,

    /// <summary>Neutral document styling that resembles a clean Word document.</summary>
    WordLike,

    /// <summary>Polished technical-document styling for guides, READMEs, and specifications.</summary>
    TechnicalDocument,

    /// <summary>GitHub-inspired styling for README-like documents.</summary>
    GitHubLike,

    /// <summary>Compact styling for dense notes and technical references.</summary>
    Compact,

    /// <summary>Report-oriented styling with stronger tables and section hierarchy.</summary>
    Report
}
