using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.EmailStore;

/// <summary>Adds email-store ingestion to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderEmailStoreExtensions {
    /// <summary>Stable handler identifier for email-store adapter registration.</summary>
    public const string HandlerId = "officeimo.reader.emailstore";

    /// <summary>Adds PST, OST, OLM, and EMLX ingestion to an isolated Reader builder.</summary>
    public static OfficeDocumentReaderBuilder AddEmailStoreHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderEmailStoreOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderEmailStoreOptions registeredOptions = ReaderEmailStoreOptionsCloner.CloneOrDefault(options);
        EmailStoreReaderOptions storeOptions = registeredOptions.StoreOptions ?? EmailStoreReaderOptions.Default;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "Email Store Reader Adapter",
            Description = "PST, OST, OLM, and EMLX adapter using the dependency-free OfficeIMO.Email.Store parser.",
            Kind = ReaderInputKind.Email,
            Extensions = new[] { ".pst", ".ost", ".olm", ".emlx" },
            ReadPath = (path, readerOptions, cancellationToken) => EmailStoreReaderAdapter.Read(
                path, readerOptions, registeredOptions, cancellationToken),
            ReadStream = (stream, sourceName, readerOptions, cancellationToken) => EmailStoreReaderAdapter.Read(
                stream, sourceName, readerOptions, registeredOptions, cancellationToken),
            ReadDocumentPath = (path, readerOptions, cancellationToken) => EmailStoreReaderAdapter.ReadDocument(
                path, readerOptions, registeredOptions, cancellationToken),
            ReadDocumentStream = (stream, sourceName, readerOptions, cancellationToken) => EmailStoreReaderAdapter.ReadDocument(
                stream, sourceName, readerOptions, registeredOptions, cancellationToken),
            DefaultMaxInputBytes = storeOptions.MaxInputBytes,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
