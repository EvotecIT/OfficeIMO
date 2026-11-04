namespace OfficeIMO.Markdown;

/// <summary>TOC scoping modes.</summary>
public enum MarkdownTocScope {
    /// <summary>Include headings from the entire document.</summary>
    Document,
    /// <summary>Include headings under the nearest preceding heading.</summary>
    PreviousHeading,
    /// <summary>Include headings under the heading matching <see cref="MarkdownTocOptions.ScopeHeadingTitle"/>.</summary>
    HeadingTitle
}

