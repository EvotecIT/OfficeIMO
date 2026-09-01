using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>
/// Root PDF lifecycle and capability container.
/// Author content through <see cref="Create(System.Action{PdfCompose}, PdfOptions)"/> or <see cref="Compose"/>;
/// use the focused capability properties for reading and existing-document operations.
/// </summary>
public sealed partial class PdfDocument {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly PdfOptions _options;
    private readonly System.Collections.Generic.Stack<System.Action<IPdfBlock>> _blockScopes;
    private readonly PdfDocumentSource? _source;
    private readonly PdfPipelineReport _pipeline;
    private readonly PdfDocumentReader _reader;

    // Metadata
    private string? _title;
    private string? _author;
    private string? _subject;
    private string? _keywords;

    private PdfDocument(PdfOptions? options = null) {
        _options = options?.Clone() ?? new PdfOptions();
        _options.MaterializeAutomaticPdfXProductionMetadata();
        _pipeline = PdfPipelineReport.Created();
        _blockScopes = new System.Collections.Generic.Stack<System.Action<IPdfBlock>>();
        _blockScopes.Push(_blocks.Add);
        Pages = new PdfDocumentPages(this);
        _reader = new PdfDocumentReader(this);
        Render = new PdfDocumentRenderer(this);
        Resources = new PdfDocumentResources(this);
        Ocr = new PdfDocumentOcr(this);
        Text = new PdfDocumentTextEditor(this);
        Images = new PdfDocumentImageEditor(this);
        Stamp = new PdfDocumentStamper(this);
        Forms = new PdfDocumentForms(this);
        Attachments = new PdfDocumentAttachments(this);
        Bookmarks = new PdfDocumentBookmarks(this);
        Annotations = new PdfDocumentAnnotations(this);
        JavaScript = new PdfDocumentJavaScript(this);
        Security = new PdfDocumentSecurity(this);
        Redactions = new PdfDocumentRedactions(this);
        Optimization = new PdfDocumentOptimization(this);
        Proof = new PdfDocumentProof(this);
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
    internal static PdfDocument Create(PdfOptions? options = null) => new PdfDocument(options);

    /// <summary>
    /// Creates and composes a PDF document through the canonical authoring DSL.
    /// </summary>
    /// <param name="compose">Document composition callback.</param>
    /// <param name="options">Optional document-wide rendering, catalog, security, and compliance options.</param>
    /// <returns>The composed <see cref="PdfDocument"/>.</returns>
    public static PdfDocument Create(System.Action<PdfCompose> compose, PdfOptions? options = null) {
        Guard.NotNull(compose, nameof(compose));
        return new PdfDocument(options).Compose(compose);
    }

    /// <summary>
    /// Loads an existing PDF from bytes and snapshots the caller-owned input once.
    /// </summary>
    public static PdfDocument Load(byte[] pdf, PdfLoadOptions? loadOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromCallerBytes(pdf, loadOptions));

    /// <summary>
    /// Opens a byte buffer owned by a trusted OfficeIMO adapter without making another snapshot.
    /// The adapter must never mutate the buffer after this call.
    /// </summary>
    internal static PdfDocument LoadOwned(byte[] pdf, PdfLoadOptions? loadOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, loadOptions));

    /// <summary>Opens an internally owned artifact together with its already validated canonical parse.</summary>
    internal static PdfDocument LoadOwned(
        byte[] pdf,
        PdfLoadOptions? loadOptions,
        PdfReadDocument readDocument) =>
        new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, loadOptions, readDocument));

    /// <summary>
    /// Loads an existing PDF from a bounded file snapshot.
    /// </summary>
    public static PdfDocument Load(string path, PdfLoadOptions? loadOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromPath(path, loadOptions));

    /// <summary>
    /// Loads a complete PDF from a readable stream. Seekable streams are read from the beginning and restored.
    /// </summary>
    public static PdfDocument Load(Stream stream, PdfLoadOptions? loadOptions = null) =>
        new PdfDocument(PdfDocumentSource.FromStream(stream, loadOptions));

    /// <summary>Asynchronously loads an existing PDF from a bounded file snapshot.</summary>
    public static async Task<PdfDocument> LoadAsync(
        string path,
        PdfLoadOptions? loadOptions = null,
        CancellationToken cancellationToken = default) {
        PdfDocumentSource source = await PdfDocumentSource
            .FromPathAsync(path, loadOptions, cancellationToken)
            .ConfigureAwait(false);
        return new PdfDocument(source);
    }

    /// <summary>Asynchronously loads a complete PDF from a readable caller-owned stream.</summary>
    public static async Task<PdfDocument> LoadAsync(
        Stream stream,
        PdfLoadOptions? loadOptions = null,
        CancellationToken cancellationToken = default) {
        PdfDocumentSource source = await PdfDocumentSource
            .FromStreamAsync(stream, loadOptions, cancellationToken)
            .ConfigureAwait(false);
        return new PdfDocument(source);
    }

    /// <summary>
    /// Page editing and extraction operations for this PDF.
    /// </summary>
    public PdfDocumentPages Pages { get; }

    /// <summary>
    /// Builds the canonical semantic document result using the structured profile by default.
    /// </summary>
    public PdfDocumentReadResult Read(PdfReadOptions? options = null, CancellationToken cancellationToken = default) {
        return PdfDocumentReadEngine.Read(this, PdfReadOptions.Resolve(options), cancellationToken);
    }

    /// <summary>
    /// Transitional internal access to focused read operations while they move to their canonical capability owners.
    /// This property is not part of the 3.3 public API.
    /// </summary>
    internal PdfDocumentReader Reader => _reader;

    /// <summary>Managed page rendering, drawing projection, and renderer diagnostics.</summary>
    public PdfDocumentRenderer Render { get; }

    /// <summary>Bounded font and raw object-resource inspection.</summary>
    public PdfDocumentResources Resources { get; }

    /// <summary>Caller-provider OCR enrichment over the canonical logical result.</summary>
    public PdfDocumentOcr Ocr { get; }

    /// <summary>Existing-page text search and editing operations.</summary>
    public PdfDocumentTextEditor Text { get; }

    /// <summary>Existing-page image placement discovery and editing operations.</summary>
    public PdfDocumentImageEditor Images { get; }

    /// <summary>Existing-document embedded and associated file editing operations.</summary>
    public PdfDocumentAttachments Attachments { get; }

    /// <summary>Existing-document bookmark editing operations.</summary>
    public PdfDocumentBookmarks Bookmarks { get; }

    /// <summary>Existing-document annotation editing operations.</summary>
    public PdfDocumentAnnotations Annotations { get; }

    /// <summary>Explicit active-content operations for named document-level JavaScript.</summary>
    public PdfDocumentJavaScript JavaScript { get; }

    /// <summary>Password encryption and digital-signature operations for this PDF.</summary>
    public PdfDocumentSecurity Security { get; }

    /// <summary>Search, planning, application, and verification operations for permanent redaction.</summary>
    public PdfDocumentRedactions Redactions { get; }

    /// <summary>Lossless optimization analysis and rewrite operations for this PDF.</summary>
    public PdfDocumentOptimization Optimization { get; }

    /// <summary>Visual and structural preservation proof operations for this PDF.</summary>
    public PdfDocumentProof Proof { get; }

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

    internal PdfReadDocument GetReadDocument(
        PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (_source is not null) {
            return _source.Read(options, cancellationToken);
        }

        return PdfReadDocument.Open(RenderBytesCore(cancellationToken), options, cancellationToken);
    }

    /// <summary>Returns a lazy canonical-parse factory only when this instance owns opened bytes.</summary>
    internal Func<PdfReadDocument>? GetOpenedReadDocumentFactory() {
        PdfDocumentSource? source = _source;
        return source is null ? null : () => source.Read();
    }

    /// <summary>
    /// Captures one byte snapshot and its canonical parse for a compound read operation.
    /// Generated documents are rendered once for the complete operation.
    /// </summary>
    internal (byte[] Bytes, PdfReadDocument Document, PdfLoadOptions Options) GetReadSnapshot(
        PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options ?? ReadOptions);
        if (_source is not null) {
            return (_source.Bytes, _source.Read(effectiveOptions, cancellationToken), effectiveOptions);
        }

        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = RenderBytesCore(cancellationToken);
        return (bytes, PdfReadDocument.Open(bytes, effectiveOptions, cancellationToken), effectiveOptions);
    }

    internal PdfLoadOptions ReadOptions {
        get {
            if (_source is not null) {
                return _source.Options;
            }

            PdfStandardEncryptionOptions? encryption = _options.EncryptionSnapshot;
            return encryption is null
                ? PdfLoadOptions.Default
                : new PdfLoadOptions {
                    Password = encryption.UserPassword,
                    AesCryptographyProvider = encryption.AesCryptographyProvider
                };
        }
    }

    internal static PdfDocument FromBytes(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, null));
    }

    internal static PdfDocument FromBytes(byte[] pdf, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfDocument(PdfDocumentSource.FromOwnedBytes(pdf, readOptions));
    }

    /// <summary>
    /// Adopts an internal operation result while carrying the source document's read contract forward.
    /// </summary>
    internal PdfDocument ApplyMutation(
        Func<byte[], byte[]> mutation,
        PdfLoadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(mutation, nameof(mutation));
        byte[] inputBytes = GetBytesForOperation();
        byte[] outputBytes = mutation(inputBytes);
        return WithBytes(inputBytes, outputBytes, readOptions, operationName);
    }

    internal PdfDocument WithBytes(
        byte[] inputBytes,
        byte[] pdf,
        PdfLoadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(inputBytes, nameof(inputBytes));
        PdfArtifactSnapshot input = _pipeline.Output ?? PdfArtifactSnapshot.Capture(inputBytes, ReadOptions);
        return WithBytes(inputBytes, input, pdf, readOptions, operationName);
    }

    internal PdfDocument WithBytes(
        byte[] inputBytes,
        PdfArtifactSnapshot input,
        byte[] pdf,
        PdfLoadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(inputBytes, nameof(inputBytes));
        Guard.NotNull(input, nameof(input));
        Guard.NotNull(pdf, nameof(pdf));
        PdfLoadOptions effectiveReadOptions = PdfLoadOptions.WithMinimumInputBytes(
            readOptions ?? ReadOptions,
            pdf.LongLength);
        PdfArtifactSnapshot output = PdfArtifactSnapshot.Capture(pdf, effectiveReadOptions);
        return WithBytes(inputBytes, input, pdf, output, effectiveReadOptions, operationName);
    }

    /// <summary>
    /// Adopts an internal operation result after reading it back and verifying the expected page count.
    /// The validated parse becomes the output document's canonical parse.
    /// </summary>
    internal PdfDocument WithBytesKnownPageCount(
        byte[] inputBytes,
        PdfArtifactSnapshot input,
        byte[] pdf,
        int outputPageCount,
        PdfLoadOptions? readOptions = null,
        [System.Runtime.CompilerServices.CallerMemberName] string operationName = "") {
        Guard.NotNull(inputBytes, nameof(inputBytes));
        Guard.NotNull(input, nameof(input));
        Guard.NotNull(pdf, nameof(pdf));
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegative(outputPageCount);
#else
        if (outputPageCount < 0) {
            throw new ArgumentOutOfRangeException(nameof(outputPageCount));
        }
#endif

        PdfLoadOptions effectiveReadOptions = PdfLoadOptions.WithMinimumInputBytes(
            readOptions ?? ReadOptions,
            pdf.LongLength);
        PdfReadDocument readback = PdfReadDocument.Open(pdf, effectiveReadOptions);
        int actualPageCount = readback.Pages.Count;
        if (actualPageCount != outputPageCount) {
            throw new InvalidOperationException("PDF operation post-save validation failed: output page count did not match the planned page count.");
        }

        PdfArtifactSnapshot output = PdfArtifactSnapshot.CaptureKnownPageCount(pdf, actualPageCount);
        return WithBytes(inputBytes, input, pdf, output, effectiveReadOptions, operationName, readback);
    }

    private PdfDocument WithBytes(
        byte[] inputBytes,
        PdfArtifactSnapshot input,
        byte[] pdf,
        PdfArtifactSnapshot output,
        PdfLoadOptions effectiveReadOptions,
        string operationName,
        PdfReadDocument? readDocument = null) {
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
        var source = readDocument is null
            ? PdfDocumentSource.FromOwnedBytes(pdf, effectiveReadOptions)
            : PdfDocumentSource.FromOwnedBytes(pdf, effectiveReadOptions, readDocument);
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
