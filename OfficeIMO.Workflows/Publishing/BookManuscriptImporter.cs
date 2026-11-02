using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Workflows;

/// <summary>Composes the canonical Word, Markdown, HTML and EPUB owners for reflowable manuscript publishing.</summary>
public static partial class BookManuscriptImporter {
    /// <summary>Imports a Word document using its semantic HTML projection and retains the projection's fidelity report.</summary>
    public static async Task<EpubManuscriptResult> ImportWordAsync(WordDocument document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options = options.Clone();
        options.Creator ??= document.BuiltinDocumentProperties.Creator;
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
        EpubManuscriptResult result = await EpubManuscript.ImportHtmlAsync(manuscript, options, cancellationToken).ConfigureAwait(false);
        return result.WithSourceReports(html.Report);
    }

    /// <summary>Imports typed Markdown as static semantic HTML, with no script or network execution.</summary>
    public static Task<EpubManuscriptResult> ImportMarkdownAsync(MarkdownDocument document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options = options.Clone();
        options.Title ??= document.FindFrontMatterEntry("title")?.Value as string;
        options.Language ??= document.FindFrontMatterEntry("language")?.Value as string ?? document.FindFrontMatterEntry("lang")?.Value as string;
        options.Creator ??= document.FindFrontMatterEntry("author")?.Value as string;
        cancellationToken.ThrowIfCancellationRequested();
        var markdownOptions = new MarkdownHtmlExportOptions {
            Kind = MarkdownHtmlKind.Document,
            Title = options.Title ?? "Book", Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None,
            AssetMode = MarkdownAssetMode.Offline, BodyClass = null, RawHtmlHandling = MarkdownRawHtmlHandling.Allow
        };
        string html = document.ToHtml(markdownOptions);
        HtmlConversionDocument manuscript = HtmlConversionDocument.Parse(html, htmlOptions, cancellationToken);
        return EpubManuscript.ImportHtmlAsync(manuscript, options, cancellationToken);
    }
}
