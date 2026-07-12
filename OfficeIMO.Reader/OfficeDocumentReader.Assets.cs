using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Reads a file and returns assets attached to its rich document result.</summary>
    public IReadOnlyList<OfficeDocumentAsset> ReadAssets(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractAssets(ReadDocument(path, options, cancellationToken), cancellationToken: cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and returns assets attached to its rich document result.</summary>
    public IReadOnlyList<OfficeDocumentAsset> ReadAssets(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractAssets(
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
        return DocumentReader.ExtractAssets(
            ReadDocument(bytes, sourceName, options, cancellationToken),
            cancellationToken: cancellationToken);
    }
}
