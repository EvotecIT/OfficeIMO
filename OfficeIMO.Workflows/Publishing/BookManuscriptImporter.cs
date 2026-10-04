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
        var wordOptions = WordToHtmlOptions.CreateSemanticDocumentProfile();
        wordOptions.Title = options.Title;
        wordOptions.Language = options.Language;
        wordOptions.MaxEmbeddedImageBytes = options.MaxResourceBytes;
        wordOptions.MaxTotalEmbeddedImageBytes = options.MaxTotalResourceBytes;
        wordOptions.ExportFootnotes = true;
        wordOptions.ExportEndnotes = true;
        HtmlTextConversionResult html = document.ToHtmlResult(wordOptions);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlConversionDocument manuscript = HtmlConversionDocument.Parse(html.Value, htmlOptions, cancellationToken);
        EpubManuscriptResult result = await EpubManuscript.ImportHtmlAsync(manuscript, options, cancellationToken).ConfigureAwait(false);
        return result.WithSourceReports(html.Report);
    }

    /// <summary>Imports typed Markdown as static semantic HTML, with no script or network execution.</summary>
    public static Task<EpubManuscriptResult> ImportMarkdownAsync(MarkdownDoc document, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options = options.Clone();
        options.Title ??= document.FindFrontMatterEntry("title")?.Value as string;
        options.Language ??= document.FindFrontMatterEntry("language")?.Value as string ?? document.FindFrontMatterEntry("lang")?.Value as string;
        options.Creator ??= document.FindFrontMatterEntry("author")?.Value as string;
        cancellationToken.ThrowIfCancellationRequested();
        var markdownOptions = new HtmlOptions {
            Title = options.Title ?? "Book", Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None,
            AssetMode = AssetMode.Offline, BodyClass = null, RawHtmlHandling = RawHtmlHandling.Allow
        };
        string html = document.ToHtmlDocument(markdownOptions);
        HtmlConversionDocument manuscript = HtmlConversionDocument.Parse(html, htmlOptions, cancellationToken);
        return EpubManuscript.ImportHtmlAsync(manuscript, options, cancellationToken);
    }
}
