using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

/// <summary>Controls Word document projection into Reader chunks.</summary>
public sealed class ReaderWordOptions {
    /// <summary>Includes footnotes in the Markdown projection.</summary>
    public bool IncludeFootnotes { get; set; } = true;

    /// <summary>
    /// When true, projects positioned text from the shared Word/PDF page layout into page blocks.
    /// The normalized block collection uses these rendered text lines; logical tables and links
    /// remain document-scoped. This is opt-in because pagination is more expensive than logical text extraction.
    /// </summary>
    public bool IncludePageLocations { get; set; }

    /// <summary>
    /// Optional font, resource, and rendering options used by shared Word pagination.
    /// LayoutOptions configures the same page layout used by PDF export.
    /// PageIndex and PageCount are ignored because Reader always builds the complete page index.
    /// </summary>
    public WordImageExportOptions? PageLocationOptions { get; set; }
}
