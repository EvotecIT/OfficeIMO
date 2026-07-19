using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Reader.Email;

/// <summary>Adds OAB ingestion to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderEmailAddressBookExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.email.address-book";

    /// <summary>Adds bounded OAB Full Details ingestion to an isolated Reader builder.</summary>
    public static OfficeDocumentReaderBuilder AddEmailAddressBookHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderEmailAddressBookOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderEmailAddressBookOptions registeredOptions =
            ReaderEmailAddressBookOptionsCloner.CloneOrDefault(options);
        OfflineAddressBookReaderOptions addressBookOptions =
            registeredOptions.AddressBookOptions ?? OfflineAddressBookReaderOptions.Default;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Email Address Book Reader Adapter",
            Description = "Bounded Outlook OAB adapter using lazy OfficeIMO.Email.AddressBook sessions.",
            Kind = ReaderInputKind.Email,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".oab" },
            ReadPath = (path, readerOptions, cancellationToken) => EmailAddressBookReaderAdapter.Read(
                path, readerOptions, registeredOptions, cancellationToken),
            ReadStream = (stream, sourceName, readerOptions, cancellationToken) => EmailAddressBookReaderAdapter.Read(
                stream, sourceName, readerOptions, registeredOptions, cancellationToken),
            ReadDocumentPath = (path, readerOptions, cancellationToken) => EmailAddressBookReaderAdapter.ReadDocument(
                path, readerOptions, registeredOptions, cancellationToken),
            ReadDocumentStream = (stream, sourceName, readerOptions, cancellationToken) => EmailAddressBookReaderAdapter.ReadDocument(
                stream, sourceName, readerOptions, registeredOptions, cancellationToken),
            DefaultMaxInputBytes = addressBookOptions.MaxInputBytes,
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
