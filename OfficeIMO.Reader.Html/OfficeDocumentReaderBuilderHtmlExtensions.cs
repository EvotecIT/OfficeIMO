namespace OfficeIMO.Reader.Html;

/// <summary>
/// Adds HTML support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderHtmlExtensions {
    /// <summary>
    /// Stable handler identifier for HTML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.html";

    /// <summary>
    /// Adds HTML ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddHtmlHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderHtmlOptions? htmlOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(htmlOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(ReaderHtmlOptions? htmlOptions) {
        ReaderHtmlOptions? registeredOptions = ReaderHtmlOptionsCloner.CloneNullable(htmlOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "HTML Reader Adapter",
            Description = "Modular HTML and MHTML adapter using OfficeIMO.Html.",
            Kind = ReaderInputKind.Html,
            Extensions = new[] { ".html", ".htm", ".xhtml", ".mht", ".mhtml" },
            ReadPath = (path, readerOptions, ct) => HtmlReaderAdapter.Read(
                htmlPath: path,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => HtmlReaderAdapter.Read(
                htmlStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => HtmlReaderAdapter.ReadDocument(
                htmlPath: path,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => HtmlReaderAdapter.ReadDocument(
                htmlStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
