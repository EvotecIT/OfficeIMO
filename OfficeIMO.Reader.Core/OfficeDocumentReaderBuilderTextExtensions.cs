namespace OfficeIMO.Reader;

/// <summary>Adds dependency-free plain-text and unknown-payload fallbacks to a Reader builder.</summary>
public static class OfficeDocumentReaderBuilderTextExtensions {
    /// <summary>Stable plain-text handler identifier.</summary>
    public const string TextHandlerId = "officeimo.reader.text";

    /// <summary>Stable last-resort unknown-payload handler identifier.</summary>
    public const string UnknownHandlerId = "officeimo.reader.unknown";

    /// <summary>
    /// Adds explicit plain-text handling and a detected-unknown fallback.
    /// </summary>
    /// <remarks>
    /// The unknown fallback has no claimed file extensions, so a registered format adapter always wins.
    /// </remarks>
    public static OfficeDocumentReaderBuilder AddPlainTextHandlers(this OfficeDocumentReaderBuilder builder, bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        builder.AddHandler(CreateRegistration(TextHandlerId, ReaderInputKind.Text, new[] { ".txt", ".log" }), replaceExisting);
        builder.AddHandler(CreateRegistration(UnknownHandlerId, ReaderInputKind.Unknown, Array.Empty<string>()), replaceExisting);
        return builder;
    }

    private static ReaderHandlerRegistration CreateRegistration(string id, ReaderInputKind kind, IReadOnlyList<string> extensions) =>
        new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = id,
            DisplayName = kind == ReaderInputKind.Text ? "Plain Text Reader" : "Unknown Payload Reader",
            Description = kind == ReaderInputKind.Text
                ? "Dependency-free bounded UTF text ingestion."
                : "Last-resort bounded byte-to-text projection for inputs no format handler accepts.",
            Kind = kind,
            Extensions = extensions,
            UseDetectedKindFallback = true,
            ReadPath = (path, options, token) => TextReaderAdapter.Read(path, kind, options, token),
            ReadStream = (stream, sourceName, options, token) => TextReaderAdapter.Read(stream, sourceName, kind, options, token),
            WarningBehavior = kind == ReaderInputKind.Unknown ? ReaderWarningBehavior.Mixed : ReaderWarningBehavior.ExceptionsOnly,
            DeterministicOutput = true
        };
}
