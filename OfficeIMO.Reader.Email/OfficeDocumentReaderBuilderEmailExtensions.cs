using OfficeIMO.Email;

namespace OfficeIMO.Reader.Email;

/// <summary>Adds direct and aggregate Email support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderEmailExtensions {
    /// <summary>Stable direct email artifact handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.email";

    /// <summary>Stable Mbox/MBX mailbox handler identifier.</summary>
    public const string MailboxHandlerId = "officeimo.reader.email.mailbox";

    /// <summary>Stable standalone iCalendar handler identifier.</summary>
    public const string CalendarHandlerId = "officeimo.reader.calendar";

    /// <summary>Stable standalone vCard handler identifier.</summary>
    public const string VCardHandlerId = "officeimo.reader.vcard";

    /// <summary>Adds EML, MSG/OFT, TNEF, Mbox/MBX, iCalendar, and vCard ingestion.</summary>
    public static OfficeDocumentReaderBuilder AddEmailHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderEmailOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderEmailOptions configured = EmailArtifactReaderAdapter.Clone(options);
        builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Email Reader",
            Description = "OfficeIMO.Email projection for messages, Outlook items, and mailbox artifacts.",
            Kind = ReaderInputKind.Email,
            Extensions = new[] { ".eml", ".msg", ".oft", ".tnef" },
            ReadDocumentPath = (path, readerOptions, token) => EmailArtifactReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => EmailArtifactReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            DefaultMaxInputBytes = configured.MessageOptions!.MaxInputBytes,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
        builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = MailboxHandlerId,
            DisplayName = "Email Mailbox Reader",
            Description = "Bounded OfficeIMO.Email projection for Mbox and MBX mailbox streams.",
            Kind = ReaderInputKind.Email,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".mbox", ".mbx" },
            ReadDocumentPath = (path, readerOptions, token) => EmailArtifactReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => EmailArtifactReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            DefaultMaxInputBytes = configured.MailboxOptions!.MaxMailboxBytes,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
        builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = CalendarHandlerId,
            DisplayName = "iCalendar Reader",
            Description = "OfficeIMO.Email projection for standalone iCalendar streams.",
            Kind = ReaderInputKind.Calendar,
            Extensions = new[] { ".ics", ".ical", ".ifb", ".vcs" },
            UseDetectedKindFallback = true,
            ReadDocumentPath = (path, readerOptions, token) => EmailArtifactReaderAdapter.ReadCalendarDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => EmailArtifactReaderAdapter.ReadCalendarDocument(stream, sourceName, readerOptions, configured, token),
            DefaultMaxInputBytes = configured.ContentLineOptions!.MaxInputBytes,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
        builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = VCardHandlerId,
            DisplayName = "vCard Reader",
            Description = "OfficeIMO.Email projection for standalone vCard streams.",
            Kind = ReaderInputKind.VCard,
            Extensions = new[] { ".vcf", ".vcard" },
            UseDetectedKindFallback = true,
            ReadDocumentPath = (path, readerOptions, token) => EmailArtifactReaderAdapter.ReadVCardDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => EmailArtifactReaderAdapter.ReadVCardDocument(stream, sourceName, readerOptions, configured, token),
            DefaultMaxInputBytes = configured.ContentLineOptions!.MaxInputBytes,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
        return builder;
    }

    /// <summary>Adds all direct, store, and Offline Address Book handlers owned by OfficeIMO.Reader.Email.</summary>
    public static OfficeDocumentReaderBuilder AddEmailHandlers(
        this OfficeDocumentReaderBuilder builder,
        ReaderEmailHandlersOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        return builder
            .AddEmailHandler(options?.Artifacts, replaceExisting)
            .AddEmailStoreHandler(options?.Stores, replaceExisting)
            .AddEmailAddressBookHandler(options?.AddressBooks, replaceExisting);
    }
}
