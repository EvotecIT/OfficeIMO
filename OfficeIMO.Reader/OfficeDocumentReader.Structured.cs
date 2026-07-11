using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Extracts a bounded structured view from an already-read document.</summary>
    public OfficeDocumentStructuredExtractionResult ExtractStructured(
        OfficeDocumentReadResult document,
        OfficeDocumentStructuredExtractionOptions? options = null,
        CancellationToken cancellationToken = default) {
        return OfficeDocumentStructuredExtractor.Extract(document, options, cancellationToken);
    }

    /// <summary>Reads a file through this reader's processors and extracts a bounded structured view.</summary>
    public OfficeDocumentStructuredExtractionResult ReadStructured(
        string path,
        ReaderOptions? readerOptions = null,
        OfficeDocumentStructuredExtractionOptions? structuredOptions = null,
        CancellationToken cancellationToken = default) {
        return ExtractStructured(ReadDocument(path, readerOptions, cancellationToken), structuredOptions, cancellationToken);
    }

    /// <summary>Reads a stream through this reader's processors and extracts a bounded structured view.</summary>
    public OfficeDocumentStructuredExtractionResult ReadStructured(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        OfficeDocumentStructuredExtractionOptions? structuredOptions = null,
        CancellationToken cancellationToken = default) {
        return ExtractStructured(
            ReadDocument(stream, sourceName, readerOptions, cancellationToken),
            structuredOptions,
            cancellationToken);
    }

    /// <summary>Reads bytes through this reader's processors and extracts a bounded structured view.</summary>
    public OfficeDocumentStructuredExtractionResult ReadStructured(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        OfficeDocumentStructuredExtractionOptions? structuredOptions = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadStructured(stream, sourceName, readerOptions, structuredOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads a file through this reader's processors and extracts a bounded structured view.</summary>
    public async Task<OfficeDocumentStructuredExtractionResult> ReadStructuredAsync(
        string path,
        ReaderOptions? readerOptions = null,
        OfficeDocumentStructuredExtractionOptions? structuredOptions = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(path, readerOptions, cancellationToken).ConfigureAwait(false);
        return ExtractStructured(document, structuredOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads a stream through this reader's processors and extracts a bounded structured view.</summary>
    public async Task<OfficeDocumentStructuredExtractionResult> ReadStructuredAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        OfficeDocumentStructuredExtractionOptions? structuredOptions = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(
            stream,
            sourceName,
            readerOptions,
            cancellationToken).ConfigureAwait(false);
        return ExtractStructured(document, structuredOptions, cancellationToken);
    }

    /// <summary>Asynchronously reads bytes through this reader's processors and extracts a bounded structured view.</summary>
    public async Task<OfficeDocumentStructuredExtractionResult> ReadStructuredAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        OfficeDocumentStructuredExtractionOptions? structuredOptions = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return await ReadStructuredAsync(
            stream,
            sourceName,
            readerOptions,
            structuredOptions,
            cancellationToken).ConfigureAwait(false);
    }
}
