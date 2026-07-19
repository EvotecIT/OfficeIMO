namespace OfficeIMO.Reader.Notebook;

/// <summary>Adds Jupyter Notebook support to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderNotebookExtensions {
    /// <summary>Stable handler identifier for notebook adapter registration.</summary>
    public const string HandlerId = "officeimo.reader.notebook";

    /// <summary>Default bounded notebook size used when <see cref="ReaderOptions.MaxInputBytes"/> is not set.</summary>
    public const long DefaultMaxInputBytes = 64L * 1024L * 1024L;

    /// <summary>Adds bounded `.ipynb` ingestion to an isolated reader builder.</summary>
    public static OfficeDocumentReaderBuilder AddNotebookHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderNotebookOptions? notebookOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderNotebookOptions registered = (notebookOptions ?? new ReaderNotebookOptions()).CloneValidated();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Jupyter Notebook Reader Adapter",
            Description = "Bounded Jupyter Notebook Markdown, code, and text-output projection.",
            Kind = ReaderInputKind.Json,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".ipynb" },
            DefaultMaxInputBytes = DefaultMaxInputBytes,
            ReadPath = (path, options, cancellationToken) => NotebookReaderAdapter
                .ReadDocument(path, options, registered, cancellationToken).Chunks,
            ReadStream = (stream, sourceName, options, cancellationToken) => NotebookReaderAdapter
                .ReadDocument(stream, sourceName, options, registered, cancellationToken).Chunks,
            ReadDocumentPath = (path, options, cancellationToken) => NotebookReaderAdapter
                .ReadDocument(path, options, registered, cancellationToken),
            ReadDocumentStream = (stream, sourceName, options, cancellationToken) => NotebookReaderAdapter
                .ReadDocument(stream, sourceName, options, registered, cancellationToken)
        }, replaceExisting);
    }
}
