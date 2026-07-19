namespace OfficeIMO.Reader.AsciiDoc;

/// <summary>Adds AsciiDoc support to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderAsciiDocExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.asciidoc";

    /// <summary>Adds `.adoc`, `.asciidoc`, and `.asc` path and stream ingestion.</summary>
    public static OfficeDocumentReaderBuilder AddAsciiDocHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderAsciiDocOptions? asciiDocOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderAsciiDocOptions registered = ReaderAsciiDocOptionsCloner.Clone(asciiDocOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "AsciiDoc Reader Adapter",
            Description = "Modular AsciiDoc adapter backed by the lossless OfficeIMO.AsciiDoc engine.",
            Kind = ReaderInputKind.AsciiDoc,
            Extensions = new[] { ".adoc", ".asciidoc", ".asc" },
            ReadPath = (path, readerOptions, cancellationToken) => AsciiDocReaderAdapter.Read(
                path,
                readerOptions,
                ReaderAsciiDocOptionsCloner.Clone(registered),
                cancellationToken),
            ReadStream = (stream, sourceName, readerOptions, cancellationToken) => AsciiDocReaderAdapter.Read(
                stream,
                sourceName,
                readerOptions,
                ReaderAsciiDocOptionsCloner.Clone(registered),
                cancellationToken),
            WarningBehavior = ReaderWarningBehavior.WarningChunksOnly,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
