namespace OfficeIMO.Reader.Latex;

/// <summary>Adds modular `.tex` ingestion to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderLatexExtensions {
    /// <summary>Stable handler ID.</summary>
    public const string HandlerId = "officeimo.reader.latex";

    /// <summary>Adds conservative extension-based `.tex` handling.</summary>
    public static OfficeDocumentReaderBuilder AddLatexHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderLatexOptions? latexOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderLatexOptions registered = ReaderLatexOptionsCloner.Clone(latexOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "LaTeX Reader Adapter",
            Description = "Modular adapter backed by the non-executing OfficeIMO LaTeX profile.",
            Kind = ReaderInputKind.Latex,
            Extensions = new[] { ".tex" },
            ReadPath = (path, readerOptions, cancellationToken) => LatexReaderAdapter.Read(
                path, readerOptions, ReaderLatexOptionsCloner.Clone(registered), cancellationToken),
            ReadStream = (stream, sourceName, readerOptions, cancellationToken) => LatexReaderAdapter.Read(
                stream, sourceName, readerOptions, ReaderLatexOptionsCloner.Clone(registered), cancellationToken),
            WarningBehavior = ReaderWarningBehavior.WarningChunksOnly,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
