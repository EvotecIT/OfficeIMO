namespace OfficeIMO.Markdown;

/// <summary>
/// Named Markdown writer profiles for common output compatibility targets.
/// </summary>
public enum MarkdownOutputProfile {
    /// <summary>OfficeIMO defaults including richer OfficeIMO markdown extensions.</summary>
    OfficeIMO,

    /// <summary>CommonMark-oriented output that avoids GitHub-only markdown syntax where practical.</summary>
    CommonMark,

    /// <summary>GitHub Flavored Markdown-oriented output for README and GitHub documentation workflows.</summary>
    GitHubFlavoredMarkdown,

    /// <summary>Portable OfficeIMO subset for stricter or parity-sensitive hosts.</summary>
    Portable
}
