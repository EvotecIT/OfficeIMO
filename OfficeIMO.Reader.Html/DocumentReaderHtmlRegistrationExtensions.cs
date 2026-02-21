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
    public static void RegisterHtmlHandler(ReaderHtmlOptions? htmlOptions = null, bool replaceExisting = false) {
        var registeredOptions = Clone(htmlOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "HTML Reader Adapter",
            Description = "Modular HTML adapter using OfficeIMO.Word.Html + OfficeIMO.Word.Markdown.",
            Kind = ReaderInputKind.Unknown,
            Extensions = new[] { ".html", ".htm" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtmlFile(
                htmlPath: path,
                readerOptions: readerOptions,
                htmlOptions: Clone(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtml(
                htmlStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                htmlOptions: Clone(registeredOptions),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters HTML ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterHtmlHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHtmlOptions? Clone(ReaderHtmlOptions? options) {
        if (options == null) return null;
        return new ReaderHtmlOptions {
            HtmlToWordOptions = Clone(options.HtmlToWordOptions),
            MarkdownOptions = Clone(options.MarkdownOptions)
        };
    }

    private static OfficeIMO.Word.Html.HtmlToWordOptions? Clone(OfficeIMO.Word.Html.HtmlToWordOptions? options) {
        if (options == null) return null;
        var clone = new OfficeIMO.Word.Html.HtmlToWordOptions {
            FontFamily = options.FontFamily,
            QuotePrefix = options.QuotePrefix,
            QuoteSuffix = options.QuoteSuffix,
            DefaultPageSize = options.DefaultPageSize,
            DefaultOrientation = options.DefaultOrientation,
            IncludeListStyles = options.IncludeListStyles,
            ContinueNumbering = options.ContinueNumbering,
            SupportsHeadingNumbering = options.SupportsHeadingNumbering,
            BasePath = options.BasePath,
            NoteReferenceType = options.NoteReferenceType,
            LinkNoteUrls = options.LinkNoteUrls,
            ImageProcessing = options.ImageProcessing,
            HttpClient = options.HttpClient,
            ResourceTimeout = options.ResourceTimeout,
            RenderPreAsTable = options.RenderPreAsTable,
            TableCaptionPosition = options.TableCaptionPosition,
            SectionTagHandling = options.SectionTagHandling
        };

        foreach (var item in options.ClassStyles) {
            clone.ClassStyles[item.Key] = item.Value;
        }

        foreach (var stylesheetPath in options.StylesheetPaths) {
            clone.StylesheetPaths.Add(stylesheetPath);
        }

        foreach (var stylesheet in options.StylesheetContents) {
            clone.StylesheetContents.Add(stylesheet);
        }

        return clone;
    }

    private static OfficeIMO.Word.Markdown.WordToMarkdownOptions? Clone(OfficeIMO.Word.Markdown.WordToMarkdownOptions? options) {
        if (options == null) return null;
        return new OfficeIMO.Word.Markdown.WordToMarkdownOptions {
            FontFamily = options.FontFamily,
            EnableUnderline = options.EnableUnderline,
            EnableHighlight = options.EnableHighlight,
            ImageExportMode = options.ImageExportMode,
            ImageDirectory = options.ImageDirectory
        };
    }
}
