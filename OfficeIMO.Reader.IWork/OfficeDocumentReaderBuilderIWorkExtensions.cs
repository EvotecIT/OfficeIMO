using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

/// <summary>Adds Apple iWork ingestion to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderIWorkExtensions {
    /// <summary>Stable handler identifier for Pages, Numbers, and Keynote.</summary>
    public const string HandlerId = "officeimo.reader.iwork";

    /// <summary>Registers bounded Pages, Numbers, and Keynote reading.</summary>
    public static OfficeDocumentReaderBuilder AddIWorkHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderIWorkOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderIWorkOptions registered = (options ?? new ReaderIWorkOptions()).Clone();
        long maximumInputBytes = registered.ReadOptions?.MaximumPackageBytes
            ?? new IWorkReadOptions().MaximumPackageBytes;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Apple iWork Reader Adapter",
            Description = "Bounded Pages, Numbers, and Keynote semantic ingestion.",
            Kind = ReaderInputKind.IWork,
            Extensions = new[] { ".pages", ".numbers", ".key" },
            DefaultMaxInputBytes = maximumInputBytes,
            MaxInputBytesCeiling = maximumInputBytes,
            ExtensionValidationProbeStream = (stream, sourceName, readerOptions, token) =>
                IWorkReaderAdapter.Probe(stream, sourceName, readerOptions, registered.Clone(), token),
            ReadPath = (path, readerOptions, token) =>
                IWorkReaderAdapter.ReadDocument(path, readerOptions, registered.Clone(), token).Chunks,
            ReadStream = (stream, sourceName, readerOptions, token) =>
                IWorkReaderAdapter.ReadDocument(stream, sourceName, readerOptions,
                    registered.Clone(), token).Chunks,
            ReadDocumentPath = (path, readerOptions, token) =>
                IWorkReaderAdapter.ReadDocument(path, readerOptions, registered.Clone(), token),
            ReadDirectoryBundle = (path, readerOptions, token) =>
                IWorkReaderAdapter.ReadDocument(path, readerOptions, registered.Clone(), token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) =>
                IWorkReaderAdapter.ReadDocument(stream, sourceName, readerOptions,
                    registered.Clone(), token)
        }, replaceExisting);
    }
}
