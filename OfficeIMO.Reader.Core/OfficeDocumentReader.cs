using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Immutable, instance-scoped OfficeIMO document reader suitable for services and concurrent hosts.
/// </summary>
/// <remarks>
/// Handler routing is frozen when the reader is built. Other <see cref="OfficeDocumentReader"/> instances
/// cannot change this reader's behavior.
/// </remarks>
public sealed partial class OfficeDocumentReader {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly SemaphoreSlim _asyncGate;
    private readonly OfficeDocumentProcessingOptions _processingOptions;

    internal OfficeDocumentReader(
        ReaderHandlerRegistrySnapshot handlers,
        int maxConcurrentReads,
        OfficeDocumentProcessorPipeline processorPipeline,
        OfficeDocumentProcessingOptions processingOptions) {
        _handlers = handlers ?? throw new ArgumentNullException(nameof(handlers));
        ProcessorPipeline = processorPipeline ?? throw new ArgumentNullException(nameof(processorPipeline));
        _processingOptions = processingOptions ?? throw new ArgumentNullException(nameof(processingOptions));
        if (maxConcurrentReads < 1 || maxConcurrentReads > DocumentReaderEngine.MaximumConcurrentReads) {
            throw new ArgumentOutOfRangeException(nameof(maxConcurrentReads));
        }
        MaxConcurrentReads = maxConcurrentReads;
        _asyncGate = new SemaphoreSlim(maxConcurrentReads, maxConcurrentReads);
    }

    /// <summary>
    /// Gets an immutable reader with no configured format handlers.
    /// </summary>
    public static OfficeDocumentReader Default { get; } = new OfficeDocumentReaderBuilder().Build();

    /// <summary>
    /// Gets the maximum number of asynchronous operations allowed in flight for this reader.
    /// </summary>
    public int MaxConcurrentReads { get; }

    /// <summary>Gets this reader's immutable ordered processor pipeline.</summary>
    public OfficeDocumentProcessorPipeline ProcessorPipeline { get; }

    /// <summary>Gets the configured processor failure behavior.</summary>
    public OfficeDocumentProcessorFailureBehavior ProcessorFailureBehavior => _processingOptions.FailureBehavior;

    /// <summary>Gets the capabilities configured for this reader.</summary>
    public IReadOnlyList<ReaderHandlerCapability> GetCapabilities() {
        return DocumentReaderEngine.GetCapabilities(_handlers);
    }

    /// <summary>Gets a machine-readable capability manifest for this reader.</summary>
    public ReaderCapabilityManifest GetCapabilityManifest() {
        return DocumentReaderEngine.GetCapabilityManifest(_handlers);
    }

    /// <summary>Gets the capability manifest for this reader as JSON.</summary>
    public string GetCapabilityManifestJson(bool indented = false) {
        return ReaderCapabilityManifestJson.Serialize(GetCapabilityManifest(), indented);
    }

    internal long? GetHandlerDefaultMaxInputBytes(string? sourceName) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.ResolveHandlerDefaultMaxInputBytes(sourceName);
        }
    }

    /// <summary>
    /// Reads a supported file using this reader's frozen handler configuration.
    /// </summary>
    public IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (ProcessorPipeline.Count == 0) {
            using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
                return DocumentReaderEngine.Read(path, options, cancellationToken);
            }
        }
        return ReadDocument(path, options, cancellationToken).Chunks ?? Array.Empty<ReaderChunk>();
    }

    /// <summary>
    /// Reads a supported stream using this reader's frozen handler configuration.
    /// </summary>
    public IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (ProcessorPipeline.Count == 0) {
            using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
                return DocumentReaderEngine.Read(stream, sourceName, options, cancellationToken);
            }
        }
        return ReadDocument(stream, sourceName, options, cancellationToken).Chunks ?? Array.Empty<ReaderChunk>();
    }

    /// <summary>
    /// Reads supported bytes using this reader's frozen handler configuration.
    /// </summary>
    public IEnumerable<ReaderChunk> Read(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (ProcessorPipeline.Count == 0) {
            using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
                return DocumentReaderEngine.Read(bytes, sourceName, options, cancellationToken);
            }
        }
        return ReadDocument(bytes, sourceName, options, cancellationToken).Chunks ?? Array.Empty<ReaderChunk>();
    }

    /// <summary>
    /// Asynchronously reads a file into normalized chunks.
    /// </summary>
    public Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (ProcessorPipeline.Count == 0) {
            return ExecuteAsync(() => DocumentReaderEngine.ReadAsync(path, options, cancellationToken), cancellationToken);
        }
        return ExecuteProcessedChunksAsync(
            () => DocumentReaderEngine.ReadDocumentAsync(path, options, cancellationToken),
            options?.ComputeHashes ?? true,
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads a stream into normalized chunks. The caller retains ownership of the stream.
    /// </summary>
    public Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (ProcessorPipeline.Count == 0) {
            return ExecuteAsync(() => DocumentReaderEngine.ReadAsync(stream, sourceName, options, cancellationToken), cancellationToken);
        }
        return ExecuteProcessedChunksAsync(
            () => DocumentReaderEngine.ReadDocumentAsync(stream, sourceName, options, cancellationToken),
            options?.ComputeHashes ?? true,
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads bytes into normalized chunks.
    /// </summary>
    public Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (ProcessorPipeline.Count == 0) {
            return ExecuteAsync(() => DocumentReaderEngine.ReadAsync(bytes, sourceName, options, cancellationToken), cancellationToken);
        }
        return ExecuteProcessedChunksAsync(
            () => DocumentReaderEngine.ReadDocumentAsync(bytes, sourceName, options, cancellationToken),
            options?.ComputeHashes ?? true,
            cancellationToken);
    }

    /// <summary>
    /// Reads a file into the shared rich document envelope.
    /// </summary>
    public OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return ProcessDocumentResult(
                DocumentReaderEngine.ReadDocument(path, options, cancellationToken),
                options?.ComputeHashes ?? true,
                cancellationToken);
        }
    }

    /// <summary>
    /// Reads a stream into the shared rich document envelope.
    /// </summary>
    public OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return ProcessDocumentResult(
                DocumentReaderEngine.ReadDocument(stream, sourceName, options, cancellationToken),
                options?.ComputeHashes ?? true,
                cancellationToken);
        }
    }

    /// <summary>
    /// Reads bytes into the shared rich document envelope.
    /// </summary>
    public OfficeDocumentReadResult ReadDocument(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return ProcessDocumentResult(
                DocumentReaderEngine.ReadDocument(bytes, sourceName, options, cancellationToken),
                options?.ComputeHashes ?? true,
                cancellationToken);
        }
    }

    /// <summary>
    /// Asynchronously reads a file into the shared rich document envelope.
    /// </summary>
    public Task<OfficeDocumentReadResult> ReadDocumentAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteProcessedDocumentAsync(
            () => DocumentReaderEngine.ReadDocumentAsync(path, options, cancellationToken),
            options?.ComputeHashes ?? true,
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads a stream into the shared rich document envelope. The caller retains ownership of the stream.
    /// </summary>
    public Task<OfficeDocumentReadResult> ReadDocumentAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteProcessedDocumentAsync(
            () => DocumentReaderEngine.ReadDocumentAsync(stream, sourceName, options, cancellationToken),
            options?.ComputeHashes ?? true,
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads bytes into the shared rich document envelope.
    /// </summary>
    public Task<OfficeDocumentReadResult> ReadDocumentAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteProcessedDocumentAsync(
            () => DocumentReaderEngine.ReadDocumentAsync(bytes, sourceName, options, cancellationToken),
            options?.ComputeHashes ?? true,
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads a bounded set of files. Results retain the input path order.
    /// </summary>
    public Task<IReadOnlyList<OfficeDocumentReadResult>> ReadDocumentsAsync(
        IEnumerable<string> paths,
        ReaderOptions? options = null,
        ReaderBatchOptions? batchOptions = null,
        CancellationToken cancellationToken = default) {
        return ReaderBatchExecutor.ExecuteAsync(
            paths,
            batchOptions,
            MaxConcurrentReads,
            MaxConcurrentReads,
            (path, token) => ReadDocumentAsync(path, options, token),
            cancellationToken);
    }

    /// <summary>
    /// Reads a file into the shared rich document JSON envelope.
    /// </summary>
    public string ReadDocumentJson(
        string path,
        ReaderOptions? options = null,
        bool indented = false,
        CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(path, options, cancellationToken), indented);
    }

    /// <summary>
    /// Reads a stream into the shared rich document JSON envelope.
    /// </summary>
    public string ReadDocumentJson(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indented = false,
        CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(
            ReadDocument(stream, sourceName, options, cancellationToken),
            indented);
    }

    /// <summary>
    /// Reads bytes into the shared rich document JSON envelope.
    /// </summary>
    public string ReadDocumentJson(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indented = false,
        CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(
            ReadDocument(bytes, sourceName, options, cancellationToken),
            indented);
    }

    /// <summary>
    /// Enumerates a folder using this reader's registered extensions and handlers.
    /// </summary>
    public IEnumerable<ReaderChunk> ReadFolder(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        return Scope(DocumentReaderEngine.ReadFolder(folderPath, folderOptions, options, onProgress, cancellationToken));
    }

    /// <summary>
    /// Enumerates a folder as one source-level payload per file.
    /// </summary>
    public IEnumerable<ReaderSourceDocument> ReadFolderDocuments(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        return Scope(DocumentReaderEngine.ReadFolderDocuments(folderPath, folderOptions, options, onProgress, cancellationToken));
    }

    /// <summary>
    /// Reads a folder and returns materialized ingestion details.
    /// </summary>
    public ReaderIngestResult ReadFolderDetailed(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeChunks = true,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.ReadFolderDetailed(folderPath, folderOptions, options, includeChunks, onProgress, cancellationToken);
        }
    }

    /// <summary>
    /// Reads a file or folder and returns source-level document payloads with optional chunk shaping.
    /// </summary>
    public ReaderPathDocumentResult ReadPathDocumentsDetailed(
        string path,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeDocumentChunks = true,
        int? maxReturnedChunks = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.ReadPathDocumentsDetailed(
                path,
                folderOptions,
                options,
                includeDocumentChunks,
                maxReturnedChunks,
                onProgress,
                cancellationToken);
        }
    }

    /// <summary>
    /// Detects a reader input kind using this reader's handler configuration.
    /// </summary>
    public ReaderInputKind DetectKind(string path) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.DetectKind(path);
        }
    }

    /// <summary>
    /// Detects a file kind from extension and bounded content evidence using this reader's handlers.
    /// </summary>
    public ReaderDetectionResult Detect(string path, ReaderDetectionOptions? options = null) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.Detect(path, options);
        }
    }

    /// <summary>
    /// Detects a stream kind from source-name and bounded content evidence using this reader's handlers.
    /// </summary>
    public ReaderDetectionResult Detect(
        Stream stream,
        string? sourceName = null,
        ReaderDetectionOptions? options = null) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.Detect(stream, sourceName, options);
        }
    }

    /// <summary>
    /// Detects a byte payload kind from source-name and bounded content evidence using this reader's handlers.
    /// </summary>
    public ReaderDetectionResult Detect(
        byte[] bytes,
        string? sourceName = null,
        ReaderDetectionOptions? options = null) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return DocumentReaderEngine.Detect(bytes, sourceName, options);
        }
    }

    /// <summary>
    /// Asynchronously detects a file kind from extension and bounded content evidence using this reader's handlers.
    /// </summary>
    public Task<ReaderDetectionResult> DetectAsync(
        string path,
        ReaderDetectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReaderEngine.DetectAsync(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Asynchronously detects a stream kind from source-name and bounded content evidence using this reader's handlers.
    /// </summary>
    public Task<ReaderDetectionResult> DetectAsync(
        Stream stream,
        string? sourceName = null,
        ReaderDetectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(
            () => DocumentReaderEngine.DetectAsync(stream, sourceName, options, cancellationToken),
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously detects a byte payload kind from source-name and bounded content evidence using this reader's handlers.
    /// </summary>
    public Task<ReaderDetectionResult> DetectAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderDetectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(
            () => DocumentReaderEngine.DetectAsync(bytes, sourceName, options, cancellationToken),
            cancellationToken);
    }

    internal IEnumerable<T> Scope<T>(IEnumerable<T> source) {
        return new ReaderHandlerScopedEnumerable<T>(_handlers, source);
    }

    private async Task<T> ExecuteAsync<T>(Func<Task<T>> action, CancellationToken cancellationToken) {
        await _asyncGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
                return await action().ConfigureAwait(false);
            }
        } finally {
            _asyncGate.Release();
        }
    }
}
