using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Extracts assets from an already-read document result.</summary>
    public IReadOnlyList<OfficeDocumentAsset> ExtractAssets(
        OfficeDocumentReadResult result,
        Func<OfficeDocumentAsset, bool>? predicate = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractAssets(result, predicate, cancellationToken);
    }

    /// <summary>Reads a file and returns assets attached to its rich document result.</summary>
    public IReadOnlyList<OfficeDocumentAsset> ReadAssets(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractAssets(ReadDocument(path, options, cancellationToken), cancellationToken: cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and returns assets attached to its rich document result.</summary>
    public IReadOnlyList<OfficeDocumentAsset> ReadAssets(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractAssets(
            ReadDocument(stream, sourceName, options, cancellationToken),
            cancellationToken: cancellationToken);
    }

    /// <summary>Reads bytes and returns assets attached to their rich document result.</summary>
    public IReadOnlyList<OfficeDocumentAsset> ReadAssets(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReaderEngine.ExtractAssets(
            ReadDocument(bytes, sourceName, options, cancellationToken),
            cancellationToken: cancellationToken);
    }
}
