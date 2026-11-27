using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Archives;

/// <summary>
/// Adds archive support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderArchiveExtensions {
    /// <summary>
    /// Stable handler identifier for archive adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.archives";

    /// <summary>
    /// Adds archive ingestion to an isolated reader builder: ZIP, 7z, Microsoft Cabinet, TAR plain or compressed with gzip,
    /// Zstandard, xz or lzip, single compressed files and, when <see cref="ReaderArchiveOptions.ReadInstallerPackages"/> is
    /// set, Windows Installer packages. Nested archives of these kinds are read through the same limits.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddArchiveHandler(
        this OfficeDocumentReaderBuilder builder,
        ArchiveTraversalOptions? zipOptions = null,
        ReaderArchiveOptions? readerArchiveOptions = null,
        bool replaceExisting = false) {
        ArgumentNullException.ThrowIfNull(builder);
        ReaderHandlerRegistration registration = CreateRegistration(zipOptions, readerArchiveOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(
        ArchiveTraversalOptions? zipOptions,
        ReaderArchiveOptions? readerArchiveOptions) {
        ArchiveTraversalOptions? registeredZipOptions = Clone(zipOptions);
        ReaderArchiveOptions? registeredReaderArchiveOptions = Clone(readerArchiveOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Archive Reader Adapter",
            Description = "Modular archive adapter that reads ZIP, 7z, CAB, plain and compressed TAR, single compressed files and optionally MSI packages, including nested archives, and emits Reader chunks.",
            Kind = ReaderInputKind.Zip,
            Extensions = Extensions(registeredReaderArchiveOptions),
            ReadPath = (path, readerOptions, ct) => ArchiveReaderAdapter.Read(
                zipPath: path,
                readerOptions: readerOptions,
                zipOptions: Clone(registeredZipOptions),
                readerArchiveOptions: Clone(registeredReaderArchiveOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => ArchiveReaderAdapter.Read(
                zipStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                zipOptions: Clone(registeredZipOptions),
                readerArchiveOptions: Clone(registeredReaderArchiveOptions),
                cancellationToken: ct)
        };
    }

    private static ArchiveTraversalOptions? Clone(ArchiveTraversalOptions? options) {
        if (options == null) return null;
        return new ArchiveTraversalOptions {
            OperationBudget = options.OperationBudget, ZipLegacyCodePage = options.ZipLegacyCodePage,
            MaxEntries = options.MaxEntries,
            MaxPhysicalEntries = options.MaxPhysicalEntries,
            MaxArchiveBytes = options.MaxArchiveBytes,
            MaxWarnings = options.MaxWarnings,
            MaxDepth = options.MaxDepth,
            MaxTotalUncompressedBytes = options.MaxTotalUncompressedBytes,
            MaxEntryUncompressedBytes = options.MaxEntryUncompressedBytes,
            MaxCompressionRatio = options.MaxCompressionRatio,
            IncludeDirectoryEntries = options.IncludeDirectoryEntries,
            DeterministicOrder = options.DeterministicOrder
        };
    }

    private static ReaderArchiveOptions? Clone(ReaderArchiveOptions? options) {
        if (options == null) return null;
        return new ReaderArchiveOptions {
            ReadNestedArchives = options.ReadNestedArchives,
            MaxNestedDepth = options.MaxNestedDepth,
            MaxNestedArchiveBytes = options.MaxNestedArchiveBytes,
            ReadInstallerPackages = options.ReadInstallerPackages
        };
    }

    private static string[] Extensions(ReaderArchiveOptions? options) {
        var extensions = new List<string> { ".zip", ".7z", ".cab", ".tar", ".tgz", ".gz", ".tzst", ".zst", ".txz", ".xz", ".tlz", ".lz" };
        if (options?.ReadInstallerPackages == true) extensions.Add(".msi");
        return extensions.ToArray();
    }
}
