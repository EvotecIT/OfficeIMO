using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Token-chunks an already-read rich document and builds its hierarchy.</summary>
    public ReaderChunkHierarchyResult ChunkDocument(
        OfficeDocumentReadResult document,
        ReaderHierarchicalChunkingOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ReaderHierarchicalChunker.Chunk(document, options, cancellationToken);

    /// <summary>Reads a file through configured processors and returns hierarchical chunks.</summary>
    public ReaderChunkHierarchyResult ReadHierarchical(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) =>
        ChunkDocument(ReadDocument(path, readerOptions, cancellationToken), chunkingOptions, cancellationToken);

    /// <summary>Reads a stream through configured processors and returns hierarchical chunks.</summary>
    public ReaderChunkHierarchyResult ReadHierarchical(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) =>
        ChunkDocument(
            ReadDocument(stream, sourceName, readerOptions, cancellationToken),
            chunkingOptions,
            cancellationToken);

    /// <summary>Reads bytes through configured processors and returns hierarchical chunks.</summary>
    public ReaderChunkHierarchyResult ReadHierarchical(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadHierarchical(stream, sourceName, readerOptions, chunkingOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads a file through configured processors and returns hierarchical chunks.</summary>
    public async Task<ReaderChunkHierarchyResult> ReadHierarchicalAsync(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderHierarchicalChunkingOptions? chunkingOptions = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(path, readerOptions, cancellationToken).ConfigureAwait(false);
        return ChunkDocument(document, chunkingOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads a stream through configured processors and returns hierarchical chunks.</summary>
    public async Task<ReaderChunkHierarchyResult> ReadHierarchicalAsync(
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

    /// <summary>Asynchronously reads bytes through configured processors and returns hierarchical chunks.</summary>
    public async Task<ReaderChunkHierarchyResult> ReadHierarchicalAsync(
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
