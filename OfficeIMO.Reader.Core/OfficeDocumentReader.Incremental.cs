using System.Threading;
using System.Threading.Tasks;
#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
#endif

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Pulls chunks on demand from incremental handlers, with a materialized fallback for other handlers.</summary>
    /// <remarks>Document processors require a complete rich result. Source hashing reads the complete source first;
    /// disable ComputeHashes for earliest first-chunk delivery. Keep the file stable throughout enumeration;
    /// detected changes throw IOException, but previously delivered chunks cannot be withdrawn.
    /// Dispose the enumerator on early termination.</remarks>
    public IEnumerable<ReaderChunk> Enumerate(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        var effective = NormalizeReadOptions(options);
        if (ProcessorPipeline.Count != 0 || Directory.Exists(path) || !HasIncrementalInput(path, pathInput: true, effective)) {
            return new ReaderMaterializedScopedEnumerable<ReaderChunk>(_handlers,
                () => ProcessorPipeline.Count == 0
                    ? DocumentReaderEngine.Read(path, effective, cancellationToken)
                    : ReadDocument(path, effective, cancellationToken).Chunks, effective, cancellationToken);
        }
        return Scope(DocumentReaderEngine.EnumerateChunks(path, effective, cancellationToken), effective);
    }

    /// <summary>Pulls chunks from a stable caller-owned stream without a full snapshot when the handler supports it.</summary>
    /// <remarks>Seekable streams read from the beginning and their position is restored on disposal.
    /// Forward-only streams read from the current position. The caller must not mutate the input during enumeration.
    /// Content-first routing, document processors, and source hashing of forward-only inputs use the materialized read contract.</remarks>
    public IEnumerable<ReaderChunk> Enumerate(Stream stream, string? sourceName = null, ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        var effective = NormalizeReadOptions(options);
        if (ProcessorPipeline.Count != 0 || !HasIncrementalInput(sourceName, pathInput: false, effective) ||
            (stream is { CanSeek: false } && effective.ComputeHashes)) {
            return new ReaderMaterializedScopedEnumerable<ReaderChunk>(_handlers,
                () => ProcessorPipeline.Count == 0
                    ? DocumentReaderEngine.Read(stream, sourceName, effective, cancellationToken)
                    : ReadDocument(stream, sourceName, effective, cancellationToken).Chunks, effective, cancellationToken);
        }
        return Scope(DocumentReaderEngine.EnumerateChunks(stream, sourceName, effective, cancellationToken), effective);
    }

    private bool HasIncrementalInput(string? sourceName, bool pathInput, ReaderOptions options) {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return options.DetectionMode != ReaderDetectionMode.PreferContent &&
                DocumentReaderEngine.HasIncrementalHandler(sourceName, pathInput);
        }
    }

    /// <summary>Enumerates chunks from caller-owned bytes.</summary>
    public IEnumerable<ReaderChunk> Enumerate(byte[] bytes, string? sourceName = null, ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return EnumerateBytes(bytes, sourceName, options, cancellationToken);
    }

    private IEnumerable<ReaderChunk> EnumerateBytes(byte[] bytes, string? sourceName, ReaderOptions? options, CancellationToken cancellationToken) {
        using var stream = new MemoryStream(bytes, writable: false);
        foreach (var chunk in Enumerate(stream, sourceName, options, cancellationToken)) yield return chunk;
    }

#if NET8_0_OR_GREATER
    /// <summary>Pulls chunks asynchronously with backpressure; synchronous handlers execute on a worker.</summary>
    public async IAsyncEnumerable<ReaderChunk> EnumerateAsync(string path, ReaderOptions? options = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) {
        options = NormalizeReadOptions(options);
        bool incremental;
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            incremental = ProcessorPipeline.Count == 0 && options?.DetectionMode != ReaderDetectionMode.PreferContent && DocumentReaderEngine.HasIncrementalHandler(path, true);
        }
        if (!incremental) {
            var chunks = ProcessorPipeline.Count == 0
                ? await ExecuteAsync(() => DocumentReaderEngine.ReadAsync(path, options, cancellationToken), cancellationToken).ConfigureAwait(false)
                : (await ReadDocumentAsync(path, options, cancellationToken).ConfigureAwait(false)).Chunks;
            foreach (var chunk in chunks) { cancellationToken.ThrowIfCancellationRequested(); yield return chunk; }
            yield break;
        }
        await _asyncGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using var iterator = Enumerate(path, options, cancellationToken).GetEnumerator();
            while (await Task.Run(iterator.MoveNext, cancellationToken).ConfigureAwait(false)) yield return iterator.Current;
        } finally { _asyncGate.Release(); }
    }

    /// <summary>Pulls caller-owned stream chunks asynchronously with backpressure and early-disposal cleanup.</summary>
    public async IAsyncEnumerable<ReaderChunk> EnumerateAsync(Stream stream, string? sourceName = null,
        ReaderOptions? options = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) {
        options = NormalizeReadOptions(options);
        bool incremental;
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            incremental = ProcessorPipeline.Count == 0 && options?.DetectionMode != ReaderDetectionMode.PreferContent &&
                DocumentReaderEngine.HasIncrementalHandler(sourceName, false);
        }
        if (!incremental) {
            var chunks = ProcessorPipeline.Count == 0
                ? await ExecuteAsync(() => DocumentReaderEngine.ReadAsync(stream, sourceName, options, cancellationToken), cancellationToken).ConfigureAwait(false)
                : (await ReadDocumentAsync(stream, sourceName, options, cancellationToken).ConfigureAwait(false)).Chunks;
            foreach (var chunk in chunks) { cancellationToken.ThrowIfCancellationRequested(); yield return chunk; }
            yield break;
        }
        await _asyncGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using var iterator = Enumerate(stream, sourceName, options, cancellationToken).GetEnumerator();
            while (await Task.Run(iterator.MoveNext, cancellationToken).ConfigureAwait(false)) yield return iterator.Current;
        } finally { _asyncGate.Release(); }
    }
    /// <summary>Enumerates caller-owned bytes asynchronously with backpressure.</summary>
    public async IAsyncEnumerable<ReaderChunk> EnumerateAsync(byte[] bytes, string? sourceName = null,
        ReaderOptions? options = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        await foreach (var chunk in EnumerateAsync(stream, sourceName, options, cancellationToken).ConfigureAwait(false)) yield return chunk;
    }
#endif
}
