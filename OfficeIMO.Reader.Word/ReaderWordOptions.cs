using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

/// <summary>Controls Word document projection into Reader chunks.</summary>
public sealed class ReaderWordOptions {
    /// <summary>Includes footnotes in the Markdown projection.</summary>
    public bool IncludeFootnotes { get; set; } = true;

    /// <summary>
    /// When true, runs the dependency-free Word paginator and maps normalized body blocks to pages.
    /// This is opt-in because pagination is more expensive than logical text extraction.
    /// </summary>
    public bool IncludePageLocations { get; set; }

    /// <summary>
    /// Optional font, resource, and rendering options used by best-effort Word pagination.
    /// PageIndex and PageCount are ignored because Reader always builds the complete page index.
    /// </summary>
    public WordImageExportOptions? PageLocationOptions { get; set; }
}
