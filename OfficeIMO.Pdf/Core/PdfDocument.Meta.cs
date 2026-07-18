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
    private readonly PdfDocumentSource? _source;
    private readonly PdfPipelineReport _pipeline;

    // Metadata
    private string? _title;
    private string? _author;
    private string? _subject;
    private string? _keywords;

    private PdfDocument(PdfOptions? options = null) {
        _options = options?.Clone() ?? new PdfOptions();
        _pipeline = PdfPipelineReport.Created();
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

    private PdfDocument(PdfDocumentSource source) : this() {
        _source = source;
        _pipeline = PdfPipelineReport.Opened(source);
    }

    private PdfDocument(PdfDocumentSource source, PdfPipelineReport pipeline) : this() {
        _source = source;
        _pipeline = pipeline;
    }

    /// <summary>
    /// Creates a new, empty PDF document with optional <paramref name="options"/>.
    /// </summary>
    /// <param name="options">Page size, margins and default font options. When null, sensible defaults are used.</param>
    /// <returns>New <see cref="PdfDocument"/> instance.</returns>
    public static PdfDocument Create(PdfOptions? options = null) => new PdfDocument(options);

    /// <summary>
    /// Opens an existing PDF from bytes and snapshots the caller-owned input once.
    /// </summary>
    public static PdfDocument Open(byte[] pdf, PdfReadOptions? readOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromCallerBytes(pdf, readOptions));

    /// <summary>
    /// Opens a byte buffer owned by a trusted OfficeIMO adapter without making another snapshot.
    /// The adapter must never mutate the buffer after this call.
    /// </summary>
    internal static PdfDocument OpenOwned(byte[] pdf, PdfReadOptions? readOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, readOptions));

    /// <summary>
    /// Opens an existing PDF from a bounded file snapshot.
    /// </summary>
    public static PdfDocument Open(string path, PdfReadOptions? readOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromPath(path, readOptions));

    /// <summary>
    /// Opens a complete PDF from a readable stream. Seekable streams are read from the beginning and restored.
    /// </summary>
    public static PdfDocument Open(Stream stream, PdfReadOptions? readOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromStream(stream, readOptions));

    /// <summary>Asynchronously opens an existing PDF from a bounded file snapshot.</summary>
    public static async Task<PdfDocument> OpenAsync(
        string path,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        PdfDocumentSource source = await PdfDocumentSource
            .FromPathAsync(path, readOptions, cancellationToken)
            .ConfigureAwait(false);
        return new PdfDocument(source);
    }

    /// <summary>Asynchronously opens a complete PDF from a readable caller-owned stream.</summary>
    public static async Task<PdfDocument> OpenAsync(
        Stream stream,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        PdfDocumentSource source = await PdfDocumentSource
            .FromStreamAsync(stream, readOptions, cancellationToken)
            .ConfigureAwait(false);
        return new PdfDocument(source);
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
    /// Immutable create/open and mutation history accumulated by this document.
    /// Save and byte-generation results append their own exact output stage.
    /// </summary>
    public PdfPipelineReport Pipeline => _pipeline;

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
        if (_source is not null) {
            throw new InvalidOperationException("This PDF was opened from existing bytes and cannot accept generated document content. Use Pages, Stamp, Forms, metadata operations, or create a new PdfDocument.");
        }
    }

    internal byte[] GetBytesForOperation() => _source?.Bytes ?? RenderBytesCore();

    internal PdfReadDocument GetReadDocument(PdfReadOptions? options = null) {
        if (_source is not null) {
            return _source.Read(options);
        }

        return PdfReadDocument.Open(RenderBytesCore(), options);
    }

    /// <summary>
    /// Captures one byte snapshot and its canonical parse for a compound read operation.
    /// Generated documents are rendered once for the complete operation.
    /// </summary>
    internal (byte[] Bytes, PdfReadDocument Document, PdfReadOptions Options) GetReadSnapshot(
        PdfReadOptions? options = null) {
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options ?? ReadOptions);
        if (_source is not null) {
            return (_source.Bytes, _source.Read(effectiveOptions), effectiveOptions);
        }

        byte[] bytes = RenderBytesCore();
        return (bytes, PdfReadDocument.Open(bytes, effectiveOptions), effectiveOptions);
    }

    internal PdfReadOptions ReadOptions {
        get {
            if (_source is not null) {
                return _source.Options;
            }

            PdfStandardEncryptionOptions? encryption = _options.EncryptionSnapshot;
            return encryption is null
                ? PdfReadOptions.Default
                : new PdfReadOptions { Password = encryption.UserPassword };
        }
    }

    internal static PdfDocument FromBytes(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, null));
    }

    internal static PdfDocument FromBytes(byte[] pdf, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, readOptions));
    }

    /// <summary>
    /// Adopts an internal operation result while carrying the source document's read contract forward.
    /// </summary>
    internal PdfDocument ApplyMutation(
        Func<byte[], byte[]> mutation,
        PdfReadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(mutation, nameof(mutation));
        byte[] inputBytes = GetBytesForOperation();
        byte[] outputBytes = mutation(inputBytes);
        return WithBytes(inputBytes, outputBytes, readOptions, operationName);
    }

    internal PdfDocument WithBytes(
        byte[] inputBytes,
        byte[] pdf,
        PdfReadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(inputBytes, nameof(inputBytes));
        PdfArtifactSnapshot input = _pipeline.Output ?? PdfArtifactSnapshot.Capture(inputBytes, ReadOptions);
        return WithBytes(inputBytes, input, pdf, readOptions, operationName);
    }

    internal PdfDocument WithBytes(
        byte[] inputBytes,
        PdfArtifactSnapshot input,
        byte[] pdf,
        PdfReadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(inputBytes, nameof(inputBytes));
        Guard.NotNull(input, nameof(input));
        Guard.NotNull(pdf, nameof(pdf));
        PdfReadOptions effectiveReadOptions = readOptions ?? ReadOptions;
        PdfArtifactSnapshot output = PdfArtifactSnapshot.Capture(pdf, effectiveReadOptions);
        PdfMutationOperation? mutationOperation = ResolveMutationOperation(operationName);
        PdfMutationExecutionMode executionMode = IsAppendOnly(inputBytes, pdf)
            ? PdfMutationExecutionMode.AppendOnly
            : PdfMutationExecutionMode.FullRewrite;
        var step = new PdfPipelineStep(
            PdfPipelineStepKind.Mutation,
            NormalizeOperationName(operationName),
            succeeded: true,
            input,
            output,
            duration: null,
            mutationOperation,
            executionMode);
        var source = PdfDocumentSource.FromOwnedBytes(pdf, effectiveReadOptions);
        return new PdfDocument(source, _pipeline.Append(step));
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
