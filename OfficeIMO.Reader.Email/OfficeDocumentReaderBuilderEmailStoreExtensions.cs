using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.Email;

/// <summary>Adds email-store ingestion to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderEmailStoreExtensions {
    /// <summary>Stable handler identifier for email-store adapter registration.</summary>
    public const string HandlerId = "officeimo.reader.email.store";

    /// <summary>Adds bounded PST, OST, OLM, and EMLX ingestion to an isolated Reader builder.</summary>
    public static OfficeDocumentReaderBuilder AddEmailStoreHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderEmailStoreOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderEmailStoreOptions registeredOptions = ReaderEmailStoreOptionsCloner.CloneOrDefault(options);
        EmailStoreReaderOptions storeOptions = registeredOptions.StoreOptions ?? EmailStoreReaderOptions.Default;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Email Store Reader Adapter",
            Description = "Bounded PST, OST, OLM, and EMLX adapter using lazy OfficeIMO.Email.Store sessions.",
            Kind = ReaderInputKind.Email,
            UseDetectedKindFallback = false,
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
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
