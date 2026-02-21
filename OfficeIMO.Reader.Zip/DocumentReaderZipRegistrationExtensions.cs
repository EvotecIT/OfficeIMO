using OfficeIMO.Zip;

namespace OfficeIMO.Reader.Zip;

/// <summary>
/// Registration helpers for plugging ZIP support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderZipRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for ZIP adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.zip";

    /// <summary>
    /// Registers ZIP ingestion into <see cref="DocumentReader"/> for the <c>.zip</c> extension.
    /// </summary>
    public static void RegisterZipHandler(
        ZipTraversalOptions? zipOptions = null,
        ReaderZipOptions? readerZipOptions = null,
        bool replaceExisting = false) {
        var registeredZipOptions = Clone(zipOptions);
        var registeredReaderZipOptions = Clone(readerZipOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "ZIP Reader Adapter",
            Description = "Modular ZIP adapter that traverses archives and emits Reader chunks.",
            Kind = ReaderInputKind.Unknown,
            Extensions = new[] { ".zip" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderZipExtensions.ReadZip(
                zipPath: path,
                readerOptions: readerOptions,
                zipOptions: Clone(registeredZipOptions),
                readerZipOptions: Clone(registeredReaderZipOptions),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters ZIP ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterZipHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
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
