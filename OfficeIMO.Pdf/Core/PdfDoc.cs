using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Root PDF document container and fluent API for composing basic PDF files.
/// Mirrors the OfficeIMO.Markdown style (H1/H2/H3, paragraph) but targets PDF output.
/// </summary>
public sealed class PdfDoc {
    private readonly List<IPdfBlock> _blocks = new();
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

    /// <summary>Adds a level-1 heading.</summary>
    public PdfDoc H1(string text, PdfAlign align = PdfAlign.Left) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(text);
#else
        if (text is null) throw new ArgumentNullException(nameof(text));
#endif
        _blocks.Add(new HeadingBlock(1, text, align)); return this; }
    /// <summary>Adds a level-2 heading.</summary>
    public PdfDoc H2(string text, PdfAlign align = PdfAlign.Left) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(text);
#else
        if (text is null) throw new ArgumentNullException(nameof(text));
#endif
        _blocks.Add(new HeadingBlock(2, text, align)); return this; }
    /// <summary>Adds a level-3 heading.</summary>
    public PdfDoc H3(string text, PdfAlign align = PdfAlign.Left) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(text);
#else
        if (text is null) throw new ArgumentNullException(nameof(text));
#endif
        _blocks.Add(new HeadingBlock(3, text, align)); return this; }
    /// <summary>Adds a paragraph of text.</summary>
    public PdfDoc P(string text, PdfAlign align = PdfAlign.Left) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(text);
#else
        if (text is null) throw new ArgumentNullException(nameof(text));
#endif
        _blocks.Add(new ParagraphBlock(text, align)); return this; }
    /// <summary>Inserts a page break.</summary>
    public PdfDoc PageBreak() { _blocks.Add(new PageBreakBlock()); return this; }

    /// <summary>Adds a simple bullet list.</summary>
    public PdfDoc Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left) {
        _blocks.Add(new BulletListBlock(items, align));
        return this;
    }

    /// <summary>Adds a simple table from rows of string arrays.</summary>
    public PdfDoc Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left) {
        _blocks.Add(new TableBlock(rows, align));
        return this;
    }

    /// <summary>
    /// Renders the document into a PDF byte array in memory.
    /// </summary>
    public byte[] ToBytes() => PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);

    /// <summary>
    /// Saves the document to <paramref name="path"/>. Creates the directory if needed.
    /// </summary>
    /// <param name="path">Destination file path, e.g. "C:\\Docs\\Report.pdf".</param>
    /// <returns>This <see cref="PdfDoc"/> for chaining.</returns>
    public PdfDoc Save(string path) {
        var bytes = ToBytes();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
        File.WriteAllBytes(path, bytes);
        return this;
    }

    /// <summary>
    /// Asynchronously saves the document to <paramref name="path"/>.
    /// </summary>
    public async System.Threading.Tasks.Task SaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        var bytes = ToBytes();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
#if NET8_0_OR_GREATER
        await fs.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
        await fs.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    // Internal getters for writer
    internal IEnumerable<IPdfBlock> Blocks => _blocks;
    internal PdfOptions Options => _options;
}
