namespace OfficeIMO.Reader.AsciiDoc;

/// <summary>Registration helpers for plugging AsciiDoc support into <see cref="DocumentReader"/>.</summary>
public static class DocumentReaderAsciiDocRegistrationExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.asciidoc";

    /// <summary>Registers `.adoc`, `.asciidoc`, and `.asc` path and stream ingestion.</summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterAsciiDocHandler(ReaderAsciiDocOptions? asciiDocOptions = null, bool replaceExisting = true) {
        ReaderAsciiDocOptions registered = ReaderAsciiDocOptionsCloner.Clone(asciiDocOptions);
        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "AsciiDoc Reader Adapter",
            Description = "Modular AsciiDoc adapter backed by the lossless OfficeIMO.AsciiDoc engine.",
            Kind = ReaderInputKind.AsciiDoc,
            Extensions = new[] { ".adoc", ".asciidoc", ".asc" },
            ReadPath = (path, readerOptions, cancellationToken) => DocumentReaderAsciiDocExtensions.ReadAsciiDocFile(
                path,
                readerOptions,
                ReaderAsciiDocOptionsCloner.Clone(registered),
                cancellationToken),
            ReadStream = (stream, sourceName, readerOptions, cancellationToken) => DocumentReaderAsciiDocExtensions.ReadAsciiDoc(
                stream,
                sourceName,
                readerOptions,
                ReaderAsciiDocOptionsCloner.Clone(registered),
                cancellationToken),
            WarningBehavior = ReaderWarningBehavior.WarningChunksOnly,
            DeterministicOutput = true
        }, replaceExisting);
    }

    /// <summary>Unregisters the AsciiDoc handler.</summary>
    public static bool UnregisterAsciiDocHandler() => DocumentReader.UnregisterHandler(HandlerId);
}
