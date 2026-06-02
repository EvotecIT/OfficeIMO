namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Controls how YAML front matter is represented in the generated PDF body.
/// </summary>
public enum MarkdownPdfFrontMatterRenderMode {
    /// <summary>Do not render front matter in the PDF body; it can still feed PDF metadata and theme selection.</summary>
    Hidden = 0,
    /// <summary>Render front matter as a polished document heading block when a title is available; otherwise render a metadata table.</summary>
    DocumentHeader = 1,
    /// <summary>Render front matter as a key/value metadata table.</summary>
    Table = 2
}
