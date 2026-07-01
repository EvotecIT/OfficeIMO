namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how automatic heading identifier slugs are generated during HTML rendering.
/// </summary>
public enum MarkdownHeadingIdentifierStyle {
    /// <summary>
    /// OfficeIMO's historical ASCII-oriented slug behavior.
    /// </summary>
    OfficeIMO,

    /// <summary>
    /// Markdig's default auto-identifier behavior, including <c>section</c> fallback for headings
    /// that do not contain ASCII identifier text.
    /// </summary>
    MarkdigDefault,

    /// <summary>
    /// Markdig/GitHub-style identifiers that preserve Unicode letters and digits where possible.
    /// </summary>
    GitHub
}
