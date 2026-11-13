using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

/// <summary>Controls Word document projection into Reader chunks.</summary>
public sealed class ReaderWordOptions {
    /// <summary>Includes footnotes and endnotes in the Markdown projection.</summary>
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

    /// <summary>
    /// When true and <see cref="ReaderOptions.Projection"/> is <see cref="ReaderProjection.TextAndMetadata"/>, DOCX
    /// bodies made of paragraphs and simple tables are read in one streaming pass without building the Word document
    /// model. The result gives the same <c>GetText()</c> output and carries only text chunks: no tables, links,
    /// blocks or document metadata. Documents with headers or footers, notes, fields, content controls, drawings,
    /// revisions or merged and nested table cells, and reads with a password, inspection or package limits, use the
    /// normal projection.
    /// </summary>
    public bool StreamTextWhenPossible { get; set; }
}
