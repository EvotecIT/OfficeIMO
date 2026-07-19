using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>Token-chunks an already-read rich document and builds its hierarchy.</summary>
    public static ReaderChunkHierarchyResult ChunkDocument(
        OfficeDocumentReadResult document,
        ReaderHierarchicalChunkingOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ReaderHierarchicalChunker.Chunk(document, options, cancellationToken);

    /// <summary>Reads a file and returns bounded token-aware hierarchical chunks.</summary>
    public static ReaderChunkHierarchyResult ReadHierarchical(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) =>
        ChunkDocument(ReadDocument(path, readerOptions, cancellationToken), chunkingOptions, cancellationToken);

    /// <summary>Reads a stream and returns bounded token-aware hierarchical chunks.</summary>
    public static ReaderChunkHierarchyResult ReadHierarchical(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) =>
        ChunkDocument(
            ReadDocument(stream, sourceName, readerOptions, cancellationToken),
            chunkingOptions,
            cancellationToken);

    /// <summary>Reads bytes and returns bounded token-aware hierarchical chunks.</summary>
    public static ReaderChunkHierarchyResult ReadHierarchical(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadHierarchical(stream, sourceName, readerOptions, chunkingOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads a file and returns bounded token-aware hierarchical chunks.</summary>
    public static async Task<ReaderChunkHierarchyResult> ReadHierarchicalAsync(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(path, readerOptions, cancellationToken).ConfigureAwait(false);
        return ChunkDocument(document, chunkingOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads a stream and returns bounded token-aware hierarchical chunks.</summary>
    public static async Task<ReaderChunkHierarchyResult> ReadHierarchicalAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(
            stream,
            sourceName,
            readerOptions,
            cancellationToken).ConfigureAwait(false);
        return ChunkDocument(document, chunkingOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads bytes and returns bounded token-aware hierarchical chunks.</summary>
    public static async Task<ReaderChunkHierarchyResult> ReadHierarchicalAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return await ReadHierarchicalAsync(
            stream,
            sourceName,
            readerOptions,
            chunkingOptions,
            cancellationToken).ConfigureAwait(false);
    }
}
