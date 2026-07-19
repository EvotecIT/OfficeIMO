namespace OfficeIMO.Reader.Markdown;

/// <summary>Adds Markdown support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderMarkdownExtensions {
    /// <summary>Stable Markdown handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.markdown";

    /// <summary>Adds Markdown and MDX text ingestion.</summary>
    public static OfficeDocumentReaderBuilder AddMarkdownHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderMarkdownOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderMarkdownOptions configured = MarkdownReaderAdapter.Clone(options);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Markdown Reader",
            Description = "OfficeIMO.Markdown AST projection with headings, tables, visuals, and source spans.",
            Kind = ReaderInputKind.Markdown,
            Extensions = new[] { ".md", ".markdown", ".mdown", ".mkd", ".mdx" },
            ReadDocumentPath = (path, readerOptions, token) => MarkdownReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => MarkdownReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
