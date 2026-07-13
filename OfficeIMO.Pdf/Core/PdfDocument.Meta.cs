using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>
/// Root PDF document container and fluent API for composing basic PDF files.
/// Mirrors the OfficeIMO.Markdown style (H1/H2/H3, paragraph) but targets PDF output.
/// </summary>
public sealed partial class PdfDocument {
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
        Annotations = new PdfDocumentAnnotations(this);
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
    /// Loads an existing PDF from bytes and snapshots the input.
    /// </summary>
    public static PdfDocument Load(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(pdf);
    }

    /// <summary>
    /// Loads an existing PDF from bytes and snapshots the input.
    /// </summary>
    public static PdfDocument Load(byte[] pdf, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(pdf, readOptions);
    }

    /// <summary>
    /// Loads an existing PDF from a file path.
    /// </summary>
    public static PdfDocument Load(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Load(File.ReadAllBytes(path));
    }

    /// <summary>
    /// Loads an existing PDF from a file path.
    /// </summary>
    public static PdfDocument Load(string path, PdfReadOptions? readOptions) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Load(File.ReadAllBytes(path), readOptions);
    }

    /// <summary>
    /// Loads a complete PDF from a readable stream. Seekable streams are read from the beginning and restored.
    /// </summary>
    public static PdfDocument Load(Stream stream) {
        return Load(OfficeStreamReader.ReadAllBytes(stream));
    }

    /// <summary>
    /// Loads a complete PDF from a readable stream. Seekable streams are read from the beginning and restored.
    /// </summary>
    public static PdfDocument Load(Stream stream, PdfReadOptions? readOptions) {
        return Load(OfficeStreamReader.ReadAllBytes(stream), readOptions);
    }

    /// <summary>Asynchronously loads an existing PDF from a file path.</summary>
    public static async Task<PdfDocument> LoadAsync(
        string path,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        string fullPath = System.IO.Path.GetFullPath(path);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 81920, useAsync: true);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
        return Load(bytes, readOptions);
    }

    /// <summary>Asynchronously loads a complete PDF from a readable caller-owned stream.</summary>
    public static async Task<PdfDocument> LoadAsync(
        Stream stream,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
        return Load(bytes, readOptions);
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

    /// <summary>Existing-document annotation editing operations.</summary>
    public PdfDocumentAnnotations Annotations { get; }

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
