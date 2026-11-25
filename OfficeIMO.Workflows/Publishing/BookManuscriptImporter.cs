using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Workflows;

/// <summary>Composes the canonical Word, Markdown, HTML and EPUB owners for reflowable manuscript publishing.</summary>
public static partial class BookManuscriptImporter {
    /// <summary>Imports a Word document using its semantic HTML projection and retains the projection's fidelity report.</summary>
    public static Task<EpubManuscriptResult> ImportWordAsync(WordDocument document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) =>
        ImportWordCoreAsync(document, options, htmlOptions, fallbackTitle: null, cancellationToken);

    private static async Task<EpubManuscriptResult> ImportWordCoreAsync(WordDocument document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions, string? fallbackTitle, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(options);
        options = options.Clone();
        options.Creator ??= document.Properties.Creator;
        if (fallbackTitle != null && !string.IsNullOrWhiteSpace(document.Properties.Title))
            options.Title ??= document.Properties.Title;
        cancellationToken.ThrowIfCancellationRequested();
        var wordOptions = WordHtmlExportOptions.CreateSemanticDocumentProfile();
        wordOptions.Title = options.Title;
        wordOptions.Language = options.Language;
        wordOptions.MaxEmbeddedImageBytes = options.MaxResourceBytes;
        wordOptions.MaxTotalEmbeddedImageBytes = options.MaxTotalResourceBytes;
        wordOptions.ExportFootnotes = true;
        wordOptions.ExportEndnotes = true;
        HtmlTextConversionResult html = document.ConvertToHtmlWithEvidence(wordOptions);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlConversionDocument manuscript = HtmlConversionDocument.Parse(html.Value, htmlOptions, cancellationToken);
        EpubManuscriptResult result = await ImportPreparedAsync(manuscript, options, fallbackTitle, useHtmlTitle: false, cancellationToken).ConfigureAwait(false);
        return result.WithSourceReports(html.Report);
    }

    /// <summary>Imports typed Markdown as static semantic HTML, with no script or network execution.</summary>
    public static Task<EpubManuscriptResult> ImportMarkdownAsync(MarkdownDocument document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) =>
        ImportMarkdownCoreAsync(document, options, htmlOptions, fallbackTitle: null, cancellationToken);

    private static Task<EpubManuscriptResult> ImportMarkdownCoreAsync(MarkdownDocument document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions, string? fallbackTitle, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(options);
        options = options.Clone();
        string? metadataTitle = MarkdownFrontMatterText.Get(document, "title");
        if (fallbackTitle == null || !string.IsNullOrWhiteSpace(metadataTitle)) options.Title ??= metadataTitle;
        options.Language ??= MarkdownFrontMatterText.Get(document, "language") ?? MarkdownFrontMatterText.Get(document, "lang");
        options.Creator ??= MarkdownFrontMatterText.Get(document, "author");
        cancellationToken.ThrowIfCancellationRequested();
        var markdownOptions = new MarkdownHtmlExportOptions {
            Kind = MarkdownHtmlKind.Document,
            Title = options.Title ?? "Book", Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None,
            AssetMode = MarkdownAssetMode.Offline, BodyClass = null, RawHtmlHandling = MarkdownRawHtmlHandling.Allow
        };
        string html = document.ToHtml(markdownOptions);
        HtmlConversionDocument manuscript = HtmlConversionDocument.Parse(html, htmlOptions, cancellationToken);
        return ImportPreparedAsync(manuscript, options, fallbackTitle, useHtmlTitle: false, cancellationToken);
    }

    private static Task<EpubManuscriptResult> ImportPreparedAsync(HtmlConversionDocument manuscript, EpubManuscriptOptions options,
        string? fallbackTitle, bool useHtmlTitle, CancellationToken cancellationToken) {
        if (fallbackTitle == null) return EpubManuscript.ImportHtmlAsync(manuscript, options, cancellationToken);
        if (options.Title == null && useHtmlTitle) {
            string? title = manuscript.Document.QuerySelector("title")?.TextContent;
            if (!string.IsNullOrWhiteSpace(title)) options.Title = title!.Trim();
        }
        options.Title ??= manuscript.Document.QuerySelectorAll("h1, h2, h3, h4, h5, h6")
            .Select(heading => heading.TextContent.Trim()).FirstOrDefault(heading => heading.Length != 0) ?? fallbackTitle;
        return EpubManuscript.ImportHtmlForWriteAsync(manuscript, options, cancellationToken);
    }
}
