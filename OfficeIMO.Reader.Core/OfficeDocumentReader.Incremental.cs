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
    public IEnumerable<ReaderChunk> EnumerateChunks(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        var effective = DocumentReaderEngine.NormalizeOptions(options);
        return Scope(ProcessorPipeline.Count == 0
            ? DocumentReaderEngine.EnumerateChunks(path, effective, cancellationToken)
            : EnumerateProcessedChunks(() => ReadDocument(path, effective, cancellationToken), cancellationToken), effective);
    }

    /// <summary>Pulls chunks from a stable caller-owned stream without a full snapshot when the handler supports it.</summary>
    /// <remarks>Seekable streams read from the beginning and their position is restored on disposal.
    /// Forward-only streams read from the current position. The caller must not mutate the input during enumeration.
    /// Content-first routing, document processors, and source hashing of forward-only inputs use the materialized read contract.</remarks>
    public IEnumerable<ReaderChunk> EnumerateChunks(Stream stream, string? sourceName = null, ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        var effective = DocumentReaderEngine.NormalizeOptions(options);
        return Scope(ProcessorPipeline.Count == 0
            ? DocumentReaderEngine.EnumerateChunks(stream, sourceName, effective, cancellationToken)
            : EnumerateProcessedChunks(() => ReadDocument(stream, sourceName, effective, cancellationToken), cancellationToken), effective);
    }

    private static IEnumerable<ReaderChunk> EnumerateProcessedChunks(Func<OfficeDocumentReadResult> read, CancellationToken cancellationToken) {
        foreach (var chunk in read().Chunks) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return chunk;
        }
    }

#if NET8_0_OR_GREATER
    /// <summary>Pulls chunks asynchronously with backpressure; synchronous handlers execute on a worker.</summary>
    public async IAsyncEnumerable<ReaderChunk> EnumerateChunksAsync(string path, ReaderOptions? options = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) {
        bool incremental;
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            incremental = ProcessorPipeline.Count == 0 && options?.DetectionMode != ReaderDetectionMode.PreferContent && DocumentReaderEngine.HasIncrementalHandler(path, true);
        }
        if (!incremental) {
            var document = await ReadDocumentAsync(path, options, cancellationToken).ConfigureAwait(false);
            foreach (var chunk in document.Chunks) { cancellationToken.ThrowIfCancellationRequested(); yield return chunk; }
            yield break;
        }
        await _asyncGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using var iterator = EnumerateChunks(path, options, cancellationToken).GetEnumerator();
            while (await Task.Run(iterator.MoveNext, cancellationToken).ConfigureAwait(false)) yield return iterator.Current;
        } finally { _asyncGate.Release(); }
    }

    /// <summary>Pulls caller-owned stream chunks asynchronously with backpressure and early-disposal cleanup.</summary>
    public async IAsyncEnumerable<ReaderChunk> EnumerateChunksAsync(Stream stream, string? sourceName = null,
        ReaderOptions? options = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) {
        bool incremental;
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            incremental = ProcessorPipeline.Count == 0 && options?.DetectionMode != ReaderDetectionMode.PreferContent &&
                DocumentReaderEngine.HasIncrementalHandler(sourceName, false);
        }
        if (!incremental) {
            var document = await ReadDocumentAsync(stream, sourceName, options, cancellationToken).ConfigureAwait(false);
            foreach (var chunk in document.Chunks) { cancellationToken.ThrowIfCancellationRequested(); yield return chunk; }
            yield break;
        }
        await _asyncGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using var iterator = EnumerateChunks(stream, sourceName, options, cancellationToken).GetEnumerator();
            while (await Task.Run(iterator.MoveNext, cancellationToken).ConfigureAwait(false)) yield return iterator.Current;
        } finally { _asyncGate.Release(); }
    }
#endif
}
