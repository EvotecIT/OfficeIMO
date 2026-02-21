namespace OfficeIMO.Reader.Html;

/// <summary>
/// Registration helpers for plugging HTML support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderHtmlRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for HTML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.html";

    /// <summary>
    /// Registers HTML ingestion into <see cref="DocumentReader"/> for <c>.html</c> and <c>.htm</c>.
    /// </summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterHtmlHandler(ReaderHtmlOptions? htmlOptions = null, bool replaceExisting = false) {
        var registeredOptions = ReaderHtmlOptionsCloner.CloneNullable(htmlOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "HTML Reader Adapter",
            Description = "Modular HTML adapter using OfficeIMO.Word.Html + OfficeIMO.Word.Markdown.",
            Kind = ReaderInputKind.Unknown,
            Extensions = new[] { ".html", ".htm" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtmlFile(
                htmlPath: path,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtml(
                htmlStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters HTML ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterHtmlHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }
}
