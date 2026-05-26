namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a table cell with optional Word-like column and row spanning.
/// </summary>
public sealed class PdfTableCell {
    /// <summary>Creates a table cell with text content, optional column/row spans, and optional URI link metadata.</summary>
    public PdfTableCell(string? text, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1) {
        if (columnSpan < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(columnSpan), "Table cell column span must be at least 1.");
        }

        if (rowSpan < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(rowSpan), "Table cell row span must be at least 1.");
        }

        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        Text = text ?? string.Empty;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        LinkUri = linkUri;
        LinkContents = linkContents;
    }

    /// <summary>Cell text content.</summary>
    public string Text { get; }

    /// <summary>Number of logical columns covered by this cell.</summary>
    public int ColumnSpan { get; }

    /// <summary>Number of logical rows covered by this cell.</summary>
    public int RowSpan { get; }

    /// <summary>Optional absolute URI linked from this cell.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional PDF annotation contents metadata for the cell link.</summary>
    public string? LinkContents { get; }

    /// <summary>Creates a single-column text cell.</summary>
    public static PdfTableCell TextCell(string? text, string? linkUri = null, string? linkContents = null) => new PdfTableCell(text, linkUri: linkUri, linkContents: linkContents);

    /// <summary>Creates a cell spanning multiple logical columns.</summary>
    public static PdfTableCell Span(string? text, int columnSpan, string? linkUri = null, string? linkContents = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents);

    /// <summary>Creates a merged cell spanning logical columns and rows.</summary>
    public static PdfTableCell Merge(string? text, int columnSpan = 1, int rowSpan = 1, string? linkUri = null, string? linkContents = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents, rowSpan);

    internal PdfTableCell Clone() => new PdfTableCell(Text, ColumnSpan, LinkUri, LinkContents, RowSpan);
}
