namespace OfficeIMO.Pdf;

/// <summary>
/// Root PDF document container and fluent API for composing basic PDF files.
/// Mirrors the OfficeIMO.Markdown style (H1/H2/H3, paragraph) but targets PDF output.
/// </summary>
public sealed partial class PdfDocument : IDisposable {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly PdfOptions _options;
    private readonly System.Collections.Generic.Stack<System.Action<IPdfBlock>> _blockScopes;
    private readonly byte[]? _loadedPdf;
    private readonly PdfReadOptions? _readOptions;

    // Metadata
    private string? _title;
    private string? _author;
    private string? _subject;
    private string? _keywords;

    private PdfDocument(PdfOptions? options = null) {
        _options = options?.Clone() ?? new PdfOptions();
        _blockScopes = new System.Collections.Generic.Stack<System.Action<IPdfBlock>>();
        _blockScopes.Push(_blocks.Add);
        Pages = new PdfDocumentPages(this);
        Read = new PdfDocumentReader(this);
        Stamp = new PdfDocumentStamper(this);
        Forms = new PdfDocumentForms(this);
        Attachments = new PdfDocumentAttachments(this);
        Bookmarks = new PdfDocumentBookmarks(this);
    }

    private PdfDocument(byte[] pdf, PdfReadOptions? readOptions = null) : this() {
        _loadedPdf = (byte[])pdf.Clone();
        _readOptions = readOptions;
    }

    /// <summary>
    /// Creates a new, empty PDF document with optional <paramref name="options"/>.
    /// </summary>
    /// <param name="options">Page size, margins and default font options. When null, sensible defaults are used.</param>
    /// <returns>New <see cref="PdfDocument"/> instance.</returns>
    public static PdfDocument Create(PdfOptions? options = null) => new PdfDocument(options);

    /// <summary>
    /// Opens an existing PDF from bytes and snapshots the input.
    /// </summary>
    public static PdfDocument Open(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(pdf);
    }

    /// <summary>
    /// Opens an existing PDF from bytes and snapshots the input.
    /// </summary>
    public static PdfDocument Open(byte[] pdf, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(pdf, readOptions);
    }

    /// <summary>
    /// Opens an existing PDF from a file path.
    /// </summary>
    public static PdfDocument Open(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Open(File.ReadAllBytes(path));
    }

    /// <summary>
    /// Opens an existing PDF from a file path.
    /// </summary>
    public static PdfDocument Open(string path, PdfReadOptions? readOptions) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Open(File.ReadAllBytes(path), readOptions);
    }

    /// <summary>
    /// Opens an existing PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocument Open(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Open(buffer.ToArray());
    }

    /// <summary>
    /// Opens an existing PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocument Open(Stream stream, PdfReadOptions? readOptions) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Open(buffer.ToArray(), readOptions);
    }

    /// <summary>
    /// Page editing and extraction operations for this PDF.
    /// </summary>
    public PdfDocumentPages Pages { get; }

    /// <summary>
    /// Readback operations for this PDF.
    /// </summary>
    public PdfDocumentReader Read { get; }

    /// <summary>Existing-document embedded and associated file editing operations.</summary>
    public PdfDocumentAttachments Attachments { get; }

    /// <summary>Existing-document bookmark editing operations.</summary>
    public PdfDocumentBookmarks Bookmarks { get; }

    /// <summary>
    /// Text and image stamping operations for this PDF.
    /// </summary>
    public PdfDocumentStamper Stamp { get; }

    /// <summary>
    /// Simple AcroForm operations for this PDF.
    /// </summary>
    public PdfDocumentForms Forms { get; }

    /// <summary>
    /// Sets PDF metadata. Only values provided are updated; missing parameters keep previous values.
    /// Pass an empty string to clear a previously assigned value.
    /// </summary>
    /// <param name="title">Document title metadata.</param>
    /// <param name="author">Document author metadata.</param>
    /// <param name="subject">Document subject metadata.</param>
    /// <param name="keywords">Document keywords metadata.</param>
    /// <returns>This <see cref="PdfDocument"/> for chaining.</returns>
    public PdfDocument Meta(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        EnsureGeneratedDocument();

        if (title != null) {
            _title = title.Length == 0 ? null : title;
        }

        if (author != null) {
            _author = author.Length == 0 ? null : author;
        }

        if (subject != null) {
            _subject = subject.Length == 0 ? null : subject;
        }

        if (keywords != null) {
            _keywords = keywords.Length == 0 ? null : keywords;
        }
        return this;
    }

    // Internal getters for writer/compose
    internal System.Collections.Generic.IEnumerable<IPdfBlock> Blocks => _blocks;
    internal PdfOptions Options => _options;

    private System.Action<IPdfBlock> CurrentBlockSink => _blockScopes.Peek();

    private void AddBlock(IPdfBlock block) {
        EnsureGeneratedDocument();
        Guard.NotNull(block, nameof(block));
        CurrentBlockSink(block);
    }

    internal void AddPageBlock(PageBlock pageBlock) { Guard.NotNull(pageBlock, nameof(pageBlock)); AddBlock(pageBlock); }

    internal void AddComposedPage(System.Action<PdfPageCompose> configure) {
        EnsureGeneratedDocument();
        Guard.NotNull(configure, nameof(configure));
        var snapshot = _options.Clone();
        if (_blocks.Count > 0) {
            snapshot.ClearPageNumberStartOverride();
        }
        var block = new PageBlock(snapshot);
        using (PushBlockScope(block.AddBlock)) {
            var page = new PdfPageCompose(this, snapshot);
            configure(page);
        }
        AddPageBlock(block);
    }

    internal System.IDisposable PushBlockScope(System.Action<IPdfBlock> addBlock) {
        Guard.NotNull(addBlock, nameof(addBlock));
        _blockScopes.Push(addBlock);
        return new Scope(this);
    }

    private void PopScope() { if (_blockScopes.Count > 1) _blockScopes.Pop(); }

    private void EnsureGeneratedDocument() {
        if (_loadedPdf is not null) {
            throw new InvalidOperationException("This PDF was opened from existing bytes and cannot accept generated document content. Use Pages, Stamp, Forms, metadata operations, or create a new PdfDocument.");
        }
    }

    internal byte[] Snapshot() {
        return ToBytes();
    }

    internal PdfReadOptions? ReadOptions => _readOptions;

    internal static PdfDocument FromBytes(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(pdf);
    }

    internal static PdfDocument FromBytes(byte[] pdf, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(pdf, readOptions);
    }

    /// <inheritdoc />
    public void Dispose() {
        // No unmanaged resources are held. IDisposable keeps the document ergonomic beside stream-backed workflows.
    }

    private sealed class Scope : System.IDisposable {
        private readonly PdfDocument _doc;
        private bool _disposed;
        public Scope(PdfDocument doc) { _doc = doc; }
        public void Dispose() {
            if (_disposed) return;
            _doc.PopScope();
            _disposed = true;
        }
    }
}

