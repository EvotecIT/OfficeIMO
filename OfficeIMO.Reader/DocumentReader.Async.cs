using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal const int DefaultMaxConcurrentReads = 4;
    internal const int MaximumConcurrentReads = 64;

    /// <summary>
    /// Asynchronously reads a file into normalized chunks. Native async handlers are awaited directly;
    /// synchronous format parsing runs after the input has been read asynchronously.
    /// </summary>
    public static async Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateFilePath(path);
        ReaderOptions opt = NormalizeOptions(options);
        EnforceFileSize(path, ResolveInitialMaxInputBytes(path, opt));

        HandlerDetectionResolution resolution = await ResolvePathHandlerAsync(path, opt, cancellationToken).ConfigureAwait(false);
        if (resolution.Handler?.ReadDocumentPathAsync != null) {
            ReaderHandlerDescriptor handler = resolution.Handler;
            cancellationToken.ThrowIfCancellationRequested();
            SourceInfo source = BuildSourceInfoFromPath(path,
                opt.ComputeHashes, cancellationToken);
            OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                handler.ReadDocumentPathAsync(path, opt, cancellationToken),
                handler.Id).ConfigureAwait(false);
            return EnrichChunks(result.Chunks, source, opt.ComputeHashes, cancellationToken);
        }

        if (resolution.Handler != null && !resolution.Handler.SupportsStreamInput) {
            return await Task.Run<IReadOnlyList<ReaderChunk>>(
                () => Read(path, opt, cancellationToken).ToArray(),
                cancellationToken).ConfigureAwait(false);
        }
        using var stream = OpenAsyncReadStream(path);
        return await ReadAsync(stream, path, opt, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously reads a stream into normalized chunks. The caller retains ownership of the stream.
    /// </summary>
    public static async Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateReadableStream(stream);
        ReaderOptions opt = NormalizeOptions(options);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "memory");

        cancellationToken.ThrowIfCancellationRequested();
        Stream readStream = await ReaderInputLimits.EnsureSeekableReadStreamAsync(
            stream,
            ResolveStreamMaxInputBytes(logicalSourceName, opt,
                stream.CanSeek),
            cancellationToken).ConfigureAwait(false);
        bool ownsReadStream = !ReferenceEquals(readStream, stream);
        try {
            HandlerDetectionResolution resolution = await ResolveStreamHandlerAsync(
                readStream,
                logicalSourceName,
                opt,
                cancellationToken).ConfigureAwait(false);
            if (resolution.Handler?.ReadDocumentStreamAsync != null) {
                ReaderHandlerDescriptor handler = resolution.Handler;
                SourceInfo source = BuildSourceInfoFromStream(readStream,
                    logicalSourceName, opt.ComputeHashes, cancellationToken);
                OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                    handler.ReadDocumentStreamAsync(readStream, logicalSourceName, opt, cancellationToken),
                    handler.Id).ConfigureAwait(false);
                return EnrichChunks(result.Chunks, source, opt.ComputeHashes, cancellationToken);
            }

            return Read(readStream, logicalSourceName, opt, cancellationToken).ToArray();
        } finally {
            if (ownsReadStream) {
                readStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Asynchronously reads bytes into normalized chunks.
    /// </summary>
    public static async Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return await ReadAsync(stream, sourceName, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously reads a file into the shared rich document envelope.
    /// </summary>
    public static async Task<OfficeDocumentReadResult> ReadDocumentAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateFilePath(path);
        ReaderOptions opt = NormalizeOptions(options);
        EnforceFileSize(path, ResolveInitialMaxInputBytes(path, opt));

        HandlerDetectionResolution resolution = await ResolvePathHandlerAsync(path, opt, cancellationToken).ConfigureAwait(false);
        if (resolution.Handler?.ReadDocumentPathAsync != null) {
            ReaderHandlerDescriptor handler = resolution.Handler;
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                handler.ReadDocumentPathAsync(path, opt, cancellationToken),
                handler.Id).ConfigureAwait(false);
            return ApplyDetectionDiagnostics(result, resolution.Detection);
        }

        if (resolution.Handler != null && !resolution.Handler.SupportsStreamInput) {
            return await Task.Run(
                () => ReadDocument(path, opt, cancellationToken),
                cancellationToken).ConfigureAwait(false);
        }
        using var stream = OpenAsyncReadStream(path);
        return await ReadDocumentAsync(stream, path, opt, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously reads a stream into the shared rich document envelope. The caller retains ownership of the stream.
    /// </summary>
    public static async Task<OfficeDocumentReadResult> ReadDocumentAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateReadableStream(stream);
        ReaderOptions opt = NormalizeOptions(options);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "memory");

        cancellationToken.ThrowIfCancellationRequested();
        Stream readStream = await ReaderInputLimits.EnsureSeekableReadStreamAsync(
            stream,
            ResolveStreamMaxInputBytes(logicalSourceName, opt,
                stream.CanSeek),
            cancellationToken).ConfigureAwait(false);
        bool ownsReadStream = !ReferenceEquals(readStream, stream);
        try {
            HandlerDetectionResolution resolution = await ResolveStreamHandlerAsync(
                readStream,
                logicalSourceName,
                opt,
                cancellationToken).ConfigureAwait(false);
            if (resolution.Handler?.ReadDocumentStreamAsync != null) {
                ReaderHandlerDescriptor handler = resolution.Handler;
                OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                    handler.ReadDocumentStreamAsync(readStream, logicalSourceName, opt, cancellationToken),
                    handler.Id).ConfigureAwait(false);
                return ApplyDetectionDiagnostics(result, resolution.Detection);
            }

            return ReadDocument(readStream, logicalSourceName, opt, cancellationToken);
        } finally {
            if (ownsReadStream) {
                readStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Asynchronously reads bytes into the shared rich document envelope.
    /// </summary>
    public static async Task<OfficeDocumentReadResult> ReadDocumentAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return await ReadDocumentAsync(stream, sourceName, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously reads a bounded set of files with deterministic result ordering.
    /// </summary>
    public static Task<IReadOnlyList<OfficeDocumentReadResult>> ReadDocumentsAsync(
        IEnumerable<string> paths,
        ReaderOptions? options = null,
        ReaderBatchOptions? batchOptions = null,
        CancellationToken cancellationToken = default) {
        return ReaderBatchExecutor.ExecuteAsync(
            paths,
            batchOptions,
            DefaultMaxConcurrentReads,
            MaximumConcurrentReads,
            (path, token) => ReadDocumentAsync(path, options, token),
            cancellationToken);
    }

    private static IReadOnlyList<ReaderChunk> EnrichChunks(
        IReadOnlyList<ReaderChunk>? chunks,
        SourceInfo source,
        bool computeHashes,
        CancellationToken cancellationToken) {
        if (chunks == null || chunks.Count == 0) {
            return Array.Empty<ReaderChunk>();
        }

        var enriched = new ReaderChunk[chunks.Count];
        for (int index = 0; index < chunks.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            enriched[index] = EnrichChunk(chunks[index], source, computeHashes);
        }
        return enriched;
    }

    private static void ValidateFilePath(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (Directory.Exists(path)) {
            throw new IOException($"'{path}' is a directory. Use {nameof(ReadFolder)}(...) to ingest directories.");
        }
        if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' doesn't exist.", path);
    }

    private static void ValidateReadableStream(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
    }

    private static InvalidOperationException CreateAsyncOnlyHandlerException(string handlerId, string inputKind) {
        return new InvalidOperationException(
            $"Reader handler '{handlerId}' only provides asynchronous {inputKind} reading. Use ReadAsync(...) or ReadDocumentAsync(...).");
    }

    private static FileStream OpenAsyncReadStream(string path) => new FileStream(
        path,
        FileMode.Open,
        FileAccess.Read,
        FileShare.ReadWrite | FileShare.Delete,
        bufferSize: 64 * 1024,
        useAsync: true);

    private static async Task<OfficeDocumentReadResult> ValidateDocumentTaskAsync(
        Task<OfficeDocumentReadResult>? task,
        string handlerId) {
        if (task == null) {
            throw new InvalidOperationException($"Reader handler '{handlerId}' returned a null asynchronous operation.");
        }

        return ValidateDocumentResult(await task.ConfigureAwait(false), handlerId);
    }
}
