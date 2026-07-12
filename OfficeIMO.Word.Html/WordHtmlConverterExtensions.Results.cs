using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    /// <summary>Imports HTML into Word and returns the document plus structured conversion evidence.</summary>
    public static HtmlToWordResult ToWordDocumentResult(this string html, HtmlToWordOptions? options = null) =>
        ToWordDocumentResultAsync(html, options).GetAwaiter().GetResult();

    /// <summary>Imports a prepared shared HTML document into Word and returns structured evidence.</summary>
    public static HtmlToWordResult ToWordDocumentResult(this HtmlConversionDocument document, HtmlToWordOptions? options = null) =>
        ToWordDocumentResultAsync(document, options).GetAwaiter().GetResult();

    /// <summary>Reads HTML from a stream, imports it into Word, and returns structured evidence.</summary>
    public static HtmlToWordResult ToWordDocumentResult(this Stream htmlStream, HtmlToWordOptions? options = null) =>
        ToWordDocumentResultAsync(htmlStream, options).GetAwaiter().GetResult();

    /// <summary>Asynchronously imports HTML into Word and returns structured evidence.</summary>
    public static async Task<HtmlToWordResult> ToWordDocumentResultAsync(
        this string html,
        HtmlToWordOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToWordOptions resolved = (options ?? new HtmlToWordOptions()).Clone();
        var converter = new HtmlToWordConverter();
        WordDocument document = await converter.ConvertAsync(html, resolved, cancellationToken).ConfigureAwait(false);
        return CreateResult(document, resolved);
    }

    /// <summary>Asynchronously imports a prepared shared HTML document into Word and returns structured evidence.</summary>
    public static async Task<HtmlToWordResult> ToWordDocumentResultAsync(
        this HtmlConversionDocument document,
        HtmlToWordOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToWordOptions resolved = (options ?? CreateWordOptionsForSharedDocument(document.ProfileContract.Profile, document.Trust)).Clone();
        resolved.ConversionProfile = document.ProfileContract.Profile;
        var converter = new HtmlToWordConverter();
        WordDocument wordDocument = await converter.ConvertAsync(
            HtmlDocumentParser.CloneDocument(document.DocumentForConversion),
            resolved,
            cancellationToken).ConfigureAwait(false);
        return CreateResult(wordDocument, resolved);
    }

    /// <summary>Asynchronously reads a stream, imports it into Word, and returns structured evidence.</summary>
    public static async Task<HtmlToWordResult> ToWordDocumentResultAsync(
        this Stream htmlStream,
        HtmlToWordOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        cancellationToken.ThrowIfCancellationRequested();
        string html = await HtmlTextIO.ReadAsync(htmlStream, cancellationToken).ConfigureAwait(false);
        return await html.ToWordDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
    }

    private static HtmlToWordResult CreateResult(WordDocument document, HtmlToWordOptions options) {
        return new HtmlToWordResult(document, options.ConversionReport);
    }
}
