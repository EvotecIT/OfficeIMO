namespace OfficeIMO.Markdown;

/// <summary>
/// High-level block categories exposed by the native markdown projection.
/// </summary>
public enum MarkdownNativeBlockKind {
    /// <summary>Heading text with level and inline markdown nodes.</summary>
    Heading,

    /// <summary>Paragraph text with inline markdown nodes.</summary>
    Paragraph,

    /// <summary>Ordered or unordered list with native list items.</summary>
    List,

    /// <summary>Quoted content with nested native blocks.</summary>
    Quote,

    /// <summary>Docs-style callout/admonition with title and nested native blocks.</summary>
    Callout,

    /// <summary>Image block with source, alternate text, title, sizing, link, and caption metadata.</summary>
    Image,

    /// <summary>Fenced or indented code block.</summary>
    Code,

    /// <summary>Markdown table with structured cells.</summary>
    Table,

    /// <summary>Semantic fenced block for diagrams, charts, networks, data views, or host-defined visuals.</summary>
    Visual,

    /// <summary>HTML details/disclosure block with summary and nested native blocks.</summary>
    Details,

    /// <summary>YAML front matter entries.</summary>
    FrontMatter,

    /// <summary>Raw HTML or HTML comment block.</summary>
    Html,

    /// <summary>Any block that does not have a specialized native projection yet.</summary>
    Other
}
