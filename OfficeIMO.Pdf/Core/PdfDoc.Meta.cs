namespace OfficeIMO.Pdf;

/// <summary>
/// Root PDF document container and fluent API for composing basic PDF files.
/// Mirrors the OfficeIMO.Markdown style (H1/H2/H3, paragraph) but targets PDF output.
/// </summary>
public sealed partial class PdfDoc {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly PdfOptions _options;

    // Metadata
    private string? _title;
    private string? _author;
    private string? _subject;
    private string? _keywords;

    private PdfDoc(PdfOptions? options = null) { _options = options ?? new PdfOptions(); }

    /// <summary>
    /// Creates a new, empty PDF document with optional <paramref name="options"/>.
    /// </summary>
    /// <param name="options">Page size, margins and default font options. When null, sensible defaults are used.</param>
    /// <returns>New <see cref="PdfDoc"/> instance.</returns>
    public static PdfDoc Create(PdfOptions? options = null) => new PdfDoc(options);

    /// <summary>
    /// Sets PDF metadata. Only values provided are updated; missing parameters keep previous values.
    /// </summary>
    /// <param name="title">Document title metadata.</param>
    /// <param name="author">Document author metadata.</param>
    /// <param name="subject">Document subject metadata.</param>
    /// <param name="keywords">Document keywords metadata.</param>
    /// <returns>This <see cref="PdfDoc"/> for chaining.</returns>
    public PdfDoc Meta(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        _title = title ?? _title;
        _author = author ?? _author;
        _subject = subject ?? _subject;
        _keywords = keywords ?? _keywords;
        return this;
    }

    // Internal getters for writer/compose
    internal System.Collections.Generic.IEnumerable<IPdfBlock> Blocks => _blocks;
    internal PdfOptions Options => _options;
}

