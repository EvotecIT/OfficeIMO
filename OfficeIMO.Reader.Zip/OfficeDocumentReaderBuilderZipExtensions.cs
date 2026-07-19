using OfficeIMO.Zip;

namespace OfficeIMO.Reader.Zip;

/// <summary>
/// Adds ZIP support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderZipExtensions {
    /// <summary>
    /// Stable handler identifier for ZIP adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.zip";

    /// <summary>
    /// Adds ZIP ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddZipHandler(
        this OfficeDocumentReaderBuilder builder,
        ZipTraversalOptions? zipOptions = null,
        ReaderZipOptions? readerZipOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(zipOptions, readerZipOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(
        ZipTraversalOptions? zipOptions,
        ReaderZipOptions? readerZipOptions) {
        ZipTraversalOptions? registeredZipOptions = Clone(zipOptions);
        ReaderZipOptions? registeredReaderZipOptions = Clone(readerZipOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "ZIP Reader Adapter",
            Description = "Modular ZIP adapter that traverses archives and emits Reader chunks.",
            Kind = ReaderInputKind.Zip,
            Extensions = new[] { ".zip" },
            ReadPath = (path, readerOptions, ct) => ZipReaderAdapter.Read(
                zipPath: path,
                readerOptions: readerOptions,
                zipOptions: Clone(registeredZipOptions),
                readerZipOptions: Clone(registeredReaderZipOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => ZipReaderAdapter.Read(
                zipStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                zipOptions: Clone(registeredZipOptions),
                readerZipOptions: Clone(registeredReaderZipOptions),
                cancellationToken: ct)
        };
    }

    private static ZipTraversalOptions? Clone(ZipTraversalOptions? options) {
        if (options == null) return null;
        return new ZipTraversalOptions {
            MaxEntries = options.MaxEntries,
            MaxDepth = options.MaxDepth,
            MaxTotalUncompressedBytes = options.MaxTotalUncompressedBytes,
            MaxEntryUncompressedBytes = options.MaxEntryUncompressedBytes,
            MaxCompressionRatio = options.MaxCompressionRatio,
            IncludeDirectoryEntries = options.IncludeDirectoryEntries,
            DeterministicOrder = options.DeterministicOrder
        };
    }

    private static ReaderZipOptions? Clone(ReaderZipOptions? options) {
        if (options == null) return null;
        return new ReaderZipOptions {
            ReadNestedZipEntries = options.ReadNestedZipEntries,
            MaxNestedDepth = options.MaxNestedDepth,
            MaxNestedArchiveBytes = options.MaxNestedArchiveBytes
        };
    }
}
