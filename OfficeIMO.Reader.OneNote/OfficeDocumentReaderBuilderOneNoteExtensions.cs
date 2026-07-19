namespace OfficeIMO.Reader.OneNote;

/// <summary>Adds offline OneNote support to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderOneNoteExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.onenote";

    /// <summary>Adds native offline <c>.one</c>, <c>.onetoc2</c>, and <c>.onepkg</c> ingestion to an isolated reader builder.</summary>
    public static OfficeDocumentReaderBuilder AddOneNoteHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderOneNoteOptions? oneNoteOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderOneNoteOptions? registered = ReaderOneNoteOptionsCloner.CloneNullable(oneNoteOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "OneNote Reader Adapter",
            Description = "Native offline OneNote adapter backed by OfficeIMO.OneNote.",
            Kind = ReaderInputKind.OneNote,
            Extensions = new[] { ".one", ".onetoc2", ".onepkg" },
            DefaultMaxInputBytes = global::OfficeIMO.OneNote.OneNoteReaderOptions.DefaultMaxInputBytes,
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true,
            ReadPath = (path, options, ct) => OneNoteReaderAdapter.Read(path, options, ReaderOneNoteOptionsCloner.CloneNullable(registered), ct),
            ReadStream = (stream, name, options, ct) => OneNoteReaderAdapter.Read(stream, name, options, ReaderOneNoteOptionsCloner.CloneNullable(registered), ct),
            ReadDocumentPath = (path, options, ct) => OneNoteReaderAdapter.ReadDocument(path, options, ReaderOneNoteOptionsCloner.CloneNullable(registered), ct),
            ReadDocumentStream = (stream, name, options, ct) => OneNoteReaderAdapter.ReadDocument(stream, name, options, ReaderOneNoteOptionsCloner.CloneNullable(registered), ct)
        }, replaceExisting);
    }
}
