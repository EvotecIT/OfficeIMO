namespace OfficeIMO.Markdown;

/// <summary>
/// High-level block categories exposed by the native markdown projection.
/// </summary>
public enum MarkdownNativeBlockKind {
    /// <summary>Paragraph text with inline markdown nodes.</summary>
    Paragraph,

    /// <summary>Fenced or indented code block.</summary>
    Code,

    /// <summary>Markdown table with structured cells.</summary>
    Table,

    /// <summary>Semantic fenced block for diagrams, charts, networks, data views, or host-defined visuals.</summary>
    Visual,

    /// <summary>Any block that does not have a specialized native projection yet.</summary>
    Other
}
