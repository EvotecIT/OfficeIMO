namespace OfficeIMO.Reader.Latex;

/// <summary>Registration helpers for modular `.tex` ingestion.</summary>
public static class DocumentReaderLatexRegistrationExtensions {
    /// <summary>Stable handler ID.</summary>
    public const string HandlerId = "officeimo.reader.latex";

    /// <summary>Registers conservative extension-based `.tex` handling.</summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterLatexHandler(ReaderLatexOptions? latexOptions = null, bool replaceExisting = true) {
        ReaderLatexOptions registered = ReaderLatexOptionsCloner.Clone(latexOptions);
        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "LaTeX Reader Adapter",
            Description = "Modular adapter backed by the non-executing OfficeIMO LaTeX profile.",
            Kind = ReaderInputKind.Latex,
            Extensions = new[] { ".tex" },
            ReadPath = (path, readerOptions, cancellationToken) => DocumentReaderLatexExtensions.ReadLatexFile(
                path, readerOptions, ReaderLatexOptionsCloner.Clone(registered), cancellationToken),
            ReadStream = (stream, sourceName, readerOptions, cancellationToken) => DocumentReaderLatexExtensions.ReadLatex(
                stream, sourceName, readerOptions, ReaderLatexOptionsCloner.Clone(registered), cancellationToken),
            WarningBehavior = ReaderWarningBehavior.WarningChunksOnly,
            DeterministicOutput = true
        }, replaceExisting);
    }

    /// <summary>Unregisters the handler.</summary>
    public static bool UnregisterLatexHandler() => DocumentReader.UnregisterHandler(HandlerId);
}
