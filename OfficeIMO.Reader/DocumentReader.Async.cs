using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    internal const int DefaultMaxConcurrentReads = 4;
    internal const int MaximumConcurrentReads = 64;

    /// <summary>
    /// Asynchronously reads a file into normalized chunks. Native async handlers are awaited directly;
    /// synchronous format engines are scheduled on a worker thread.
    /// </summary>
    public static async Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateFilePath(path);
        ReaderOptions opt = NormalizeOptions(options);
        EnforceFileSize(path, opt.MaxInputBytes);

        HandlerDetectionResolution resolution = await ResolvePathHandlerAsync(path, opt, cancellationToken).ConfigureAwait(false);
        if (resolution.Handler?.ReadDocumentPathAsync != null) {
            ReaderHandlerDescriptor handler = resolution.Handler;
            cancellationToken.ThrowIfCancellationRequested();
            SourceInfo source = BuildSourceInfoFromPath(path, opt.ComputeHashes);
            OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                handler.ReadDocumentPathAsync(path, opt, cancellationToken),
                handler.Id).ConfigureAwait(false);
            return EnrichChunks(result.Chunks, source, opt.ComputeHashes, cancellationToken);
        }

        return await Task.Run<IReadOnlyList<ReaderChunk>>(
            () => Read(path, opt, cancellationToken).ToArray(),
            cancellationToken).ConfigureAwait(false);
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
            opt.MaxInputBytes,
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
                SourceInfo source = BuildSourceInfoFromStream(readStream, logicalSourceName, opt.ComputeHashes);
                OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                    handler.ReadDocumentStreamAsync(readStream, logicalSourceName, opt, cancellationToken),
                    handler.Id).ConfigureAwait(false);
                return EnrichChunks(result.Chunks, source, opt.ComputeHashes, cancellationToken);
            }

            return await Task.Run<IReadOnlyList<ReaderChunk>>(
                () => Read(readStream, logicalSourceName, opt, cancellationToken).ToArray(),
                cancellationToken).ConfigureAwait(false);
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
        EnforceFileSize(path, opt.MaxInputBytes);

        HandlerDetectionResolution resolution = await ResolvePathHandlerAsync(path, opt, cancellationToken).ConfigureAwait(false);
        if (resolution.Handler?.ReadDocumentPathAsync != null) {
            ReaderHandlerDescriptor handler = resolution.Handler;
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentReadResult result = await ValidateDocumentTaskAsync(
                handler.ReadDocumentPathAsync(path, opt, cancellationToken),
                handler.Id).ConfigureAwait(false);
            return ApplyDetectionDiagnostics(result, resolution.Detection);
        }

        return await Task.Run(
            () => ReadDocument(path, opt, cancellationToken),
            cancellationToken).ConfigureAwait(false);
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
            opt.MaxInputBytes,
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

            return await Task.Run(
                () => ReadDocument(readStream, logicalSourceName, opt, cancellationToken),
                cancellationToken).ConfigureAwait(false);
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

    private static async Task<OfficeDocumentReadResult> ValidateDocumentTaskAsync(
        Task<OfficeDocumentReadResult>? task,
        string handlerId) {
        if (task == null) {
            throw new InvalidOperationException($"Reader handler '{handlerId}' returned a null asynchronous operation.");
        }

        return ValidateDocumentResult(await task.ConfigureAwait(false), handlerId);
    }
}
