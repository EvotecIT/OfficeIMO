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
/// Handler routing is frozen when the reader is built. Static <see cref="DocumentReader"/> registrations
/// and other <see cref="OfficeDocumentReader"/> instances cannot change this reader's behavior.
/// </remarks>
public sealed class OfficeDocumentReader {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly SemaphoreSlim _asyncGate;

    internal OfficeDocumentReader(ReaderHandlerRegistrySnapshot handlers, int maxConcurrentReads) {
        _handlers = handlers ?? throw new ArgumentNullException(nameof(handlers));
        if (maxConcurrentReads < 1 || maxConcurrentReads > DocumentReader.MaximumConcurrentReads) {
            throw new ArgumentOutOfRangeException(nameof(maxConcurrentReads));
        }
        MaxConcurrentReads = maxConcurrentReads;
        _asyncGate = new SemaphoreSlim(maxConcurrentReads, maxConcurrentReads);
    }

    /// <summary>
    /// Gets a reader with built-in handlers and no custom registrations.
    /// </summary>
    public static OfficeDocumentReader Default { get; } = new OfficeDocumentReaderBuilder().Build();

    /// <summary>
    /// Gets the maximum number of asynchronous operations allowed in flight for this reader.
    /// </summary>
    public int MaxConcurrentReads { get; }

    /// <summary>
    /// Reads a supported file using this reader's frozen handler configuration.
    /// </summary>
    public IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return Scope(DocumentReader.Read(path, options, cancellationToken));
    }

    /// <summary>
    /// Reads a supported stream using this reader's frozen handler configuration.
    /// </summary>
    public IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return Scope(DocumentReader.Read(stream, sourceName, options, cancellationToken));
    }

    /// <summary>
    /// Reads supported bytes using this reader's frozen handler configuration.
    /// </summary>
    public IEnumerable<ReaderChunk> Read(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return Scope(DocumentReader.Read(bytes, sourceName, options, cancellationToken));
    }

    /// <summary>
    /// Asynchronously reads a file into normalized chunks.
    /// </summary>
    public Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReader.ReadAsync(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads a stream into normalized chunks. The caller retains ownership of the stream.
    /// </summary>
    public Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReader.ReadAsync(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads bytes into normalized chunks.
    /// </summary>
    public Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReader.ReadAsync(bytes, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a file into the shared rich document envelope.
    /// </summary>
    public OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadDocument(path, options, cancellationToken);
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadDocument(stream, sourceName, options, cancellationToken);
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadDocument(bytes, sourceName, options, cancellationToken);
        }
    }

    /// <summary>
    /// Asynchronously reads a file into the shared rich document envelope.
    /// </summary>
    public Task<OfficeDocumentReadResult> ReadDocumentAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReader.ReadDocumentAsync(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads a stream into the shared rich document envelope. The caller retains ownership of the stream.
    /// </summary>
    public Task<OfficeDocumentReadResult> ReadDocumentAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReader.ReadDocumentAsync(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads bytes into the shared rich document envelope.
    /// </summary>
    public Task<OfficeDocumentReadResult> ReadDocumentAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ExecuteAsync(() => DocumentReader.ReadDocumentAsync(bytes, sourceName, options, cancellationToken), cancellationToken);
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadDocumentJson(path, options, indented, cancellationToken);
        }
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadDocumentJson(stream, sourceName, options, indented, cancellationToken);
        }
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadDocumentJson(bytes, sourceName, options, indented, cancellationToken);
        }
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
        return Scope(DocumentReader.ReadFolder(folderPath, folderOptions, options, onProgress, cancellationToken));
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
        return Scope(DocumentReader.ReadFolderDocuments(folderPath, folderOptions, options, onProgress, cancellationToken));
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadFolderDetailed(folderPath, folderOptions, options, includeChunks, onProgress, cancellationToken);
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.ReadPathDocumentsDetailed(
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
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.DetectKind(path);
        }
    }

    /// <summary>
    /// Lists capabilities visible to this reader instance.
    /// </summary>
    public IReadOnlyList<ReaderHandlerCapability> GetCapabilities(bool includeBuiltIn = true, bool includeCustom = true) {
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.GetCapabilities(includeBuiltIn, includeCustom);
        }
    }

    /// <summary>
    /// Builds a capability manifest for this reader instance.
    /// </summary>
    public ReaderCapabilityManifest GetCapabilityManifest(bool includeBuiltIn = true, bool includeCustom = true) {
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.GetCapabilityManifest(includeBuiltIn, includeCustom);
        }
    }

    /// <summary>
    /// Builds a JSON capability manifest for this reader instance.
    /// </summary>
    public string GetCapabilityManifestJson(bool includeBuiltIn = true, bool includeCustom = true, bool indented = false) {
        using (DocumentReader.UseHandlerRegistry(_handlers)) {
            return DocumentReader.GetCapabilityManifestJson(includeBuiltIn, includeCustom, indented);
        }
    }

    private IEnumerable<T> Scope<T>(IEnumerable<T> source) {
        return new ReaderHandlerScopedEnumerable<T>(_handlers, source);
    }

    private async Task<T> ExecuteAsync<T>(Func<Task<T>> action, CancellationToken cancellationToken) {
        await _asyncGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using (DocumentReader.UseHandlerRegistry(_handlers)) {
                return await action().ConfigureAwait(false);
            }
        } finally {
            _asyncGate.Release();
        }
    }
}
