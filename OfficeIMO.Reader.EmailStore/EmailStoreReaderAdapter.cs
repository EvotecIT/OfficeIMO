using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.EmailStore;

internal static class EmailStoreReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions readerOptions,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailStoreReaderOptions effective = ReaderEmailStoreOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        EmailStoreReadResult readResult = new global::OfficeIMO.Email.Store.EmailStoreReader(effective)
            .Read(path, cancellationToken);
        EmailStoreProjection projection = EmailStoreReaderProjection.Create(readResult, path, cancellationToken);
        return DocumentReaderEngine.ProjectEmailDocumentsToChunks(
            projection.Documents,
            projection.LogicalPaths,
            projection.Diagnostics,
            projection.EmailFormat,
            path,
            readerOptions,
            cancellationToken);
    }

    internal static IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName,
        ReaderOptions readerOptions,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailStoreReaderOptions effective = ReaderEmailStoreOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        string logicalSourceName = NormalizeSourceName(sourceName);
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream, effective.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            EmailStoreReadResult readResult = new global::OfficeIMO.Email.Store.EmailStoreReader(effective)
                .Read(parseStream, logicalSourceName, cancellationToken);
            EmailStoreProjection projection = EmailStoreReaderProjection.Create(
                readResult, logicalSourceName, cancellationToken);
            return DocumentReaderEngine.ProjectEmailDocumentsToChunks(
                projection.Documents,
                projection.LogicalPaths,
                projection.Diagnostics,
                projection.EmailFormat,
                logicalSourceName,
                readerOptions,
                cancellationToken).ToArray();
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    internal static OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions readerOptions,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailStoreReaderOptions effective = ReaderEmailStoreOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        EmailStoreReadResult readResult = new global::OfficeIMO.Email.Store.EmailStoreReader(effective)
            .Read(path, cancellationToken);
        EmailStoreProjection projection = EmailStoreReaderProjection.Create(readResult, path, cancellationToken);
        OfficeDocumentReadResult result = DocumentReaderEngine.ProjectEmailDocumentsToPathResult(
            projection.Documents,
            projection.LogicalPaths,
            projection.Diagnostics,
            projection.EmailFormat,
            path,
            path,
            readerOptions,
            cancellationToken);
        return EmailStoreReaderProjection.EnrichResult(result, projection);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName,
        ReaderOptions readerOptions,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailStoreReaderOptions effective = ReaderEmailStoreOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        string logicalSourceName = NormalizeSourceName(sourceName);
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream, effective.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            EmailStoreReadResult readResult = new global::OfficeIMO.Email.Store.EmailStoreReader(effective)
                .Read(parseStream, logicalSourceName, cancellationToken);
            EmailStoreProjection projection = EmailStoreReaderProjection.Create(
                readResult, logicalSourceName, cancellationToken);
            OfficeDocumentReadResult result = DocumentReaderEngine.ProjectEmailDocumentsToStreamResult(
                projection.Documents,
                projection.LogicalPaths,
                projection.Diagnostics,
                projection.EmailFormat,
                logicalSourceName,
                parseStream,
                readerOptions,
                cancellationToken);
            return EmailStoreReaderProjection.EnrichResult(result, projection);
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    private static string NormalizeSourceName(string? sourceName) {
        return string.IsNullOrWhiteSpace(sourceName) ? "email-store.bin" : sourceName!.Trim();
    }
}
