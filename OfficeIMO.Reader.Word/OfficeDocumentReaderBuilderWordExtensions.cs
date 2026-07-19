namespace OfficeIMO.Reader.Word;

/// <summary>Adds Word support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderWordExtensions {
    /// <summary>Stable Word handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.word";

    /// <summary>Adds DOCX, DOCM, and legacy DOC ingestion.</summary>
    public static OfficeDocumentReaderBuilder AddWordHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderWordOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderWordOptions configured = WordReaderAdapter.Clone(options);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Word Reader",
            Description = "OfficeIMO.Word Markdown and structured document projection.",
            Kind = ReaderInputKind.Word,
            Extensions = new[] { ".docx", ".docm", ".doc" },
            ReadDocumentPath = (path, readerOptions, token) => WordReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => WordReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            ProbeStream = (stream, sourceName, readerOptions, token) => WordReaderAdapter.ProbeEncryptedOpenXml(stream, readerOptions, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
