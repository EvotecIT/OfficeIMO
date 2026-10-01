using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.Email;

internal static class EmailStoreReaderAdapter {
    internal static bool IncludesEmbeddedContent(ReaderEmailStoreOptions options) =>
        EmailStoreReaderProjection.GetItemReadOptions(options).Includes(EmailStoreItemReadParts.EmbeddedItems);

    internal static IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions readerOptions,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailStoreReaderOptions effective = ReaderEmailStoreOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        using (EmailStoreSession session = EmailStoreSession.Open(path, effective, cancellationToken)) {
            EmailStoreProjection projection = EmailStoreReaderProjection.Create(
                session, path, adapterOptions, cancellationToken);
            return EmailReaderProjection.ProjectEmailDocumentsToChunks(
                projection.Documents,
                projection.LogicalPaths,
                projection.Diagnostics,
                projection.EmailFormat,
                path,
                readerOptions,
                cancellationToken, includeEmbeddedMessageContent: IncludesEmbeddedContent(adapterOptions));
        }
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
            using (EmailStoreSession session = EmailStoreSession.Open(
                parseStream, logicalSourceName, effective, leaveOpen: true, cancellationToken)) {
                EmailStoreProjection projection = EmailStoreReaderProjection.Create(
                    session, logicalSourceName, adapterOptions, cancellationToken);
                return EmailReaderProjection.ProjectEmailDocumentsToChunks(
                    projection.Documents,
                    projection.LogicalPaths,
                    projection.Diagnostics,
                    projection.EmailFormat,
                    logicalSourceName,
                    readerOptions,
                    cancellationToken, includeEmbeddedMessageContent: IncludesEmbeddedContent(adapterOptions)).ToArray();
            }
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
        using (EmailStoreSession session = EmailStoreSession.Open(path, effective, cancellationToken)) {
            EmailStoreProjection projection = EmailStoreReaderProjection.Create(
                session, path, adapterOptions, cancellationToken);
            OfficeDocumentReadResult result = EmailReaderProjection.ProjectEmailDocumentsToPathResult(
                projection.Documents,
                projection.LogicalPaths,
                projection.Diagnostics,
                projection.EmailFormat,
                path,
                path,
                readerOptions,
                cancellationToken,
                computeSourceHash: adapterOptions.ComputeSourceHash, includeEmbeddedMessageContent: IncludesEmbeddedContent(adapterOptions));
            return EmailStoreReaderProjection.EnrichResult(result, projection);
        }
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
            using (EmailStoreSession session = EmailStoreSession.Open(
                parseStream, logicalSourceName, effective, leaveOpen: true, cancellationToken)) {
                EmailStoreProjection projection = EmailStoreReaderProjection.Create(
                    session, logicalSourceName, adapterOptions, cancellationToken);
                OfficeDocumentReadResult result = EmailReaderProjection.ProjectEmailDocumentsToStreamResult(
                    projection.Documents,
                    projection.LogicalPaths,
                    projection.Diagnostics,
                    projection.EmailFormat,
                    logicalSourceName,
                    parseStream,
                    readerOptions,
                    cancellationToken,
                    computeSourceHash: adapterOptions.ComputeSourceHash, includeEmbeddedMessageContent: IncludesEmbeddedContent(adapterOptions));
                return EmailStoreReaderProjection.EnrichResult(result, projection);
            }
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    private static string NormalizeSourceName(string? sourceName) {
        return string.IsNullOrWhiteSpace(sourceName) ? "email-store.bin" : sourceName!.Trim();
    }
}
