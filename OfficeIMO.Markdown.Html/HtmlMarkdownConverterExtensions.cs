using OfficeIMO.Markdown;
using OfficeIMO.Html;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Extension helpers for converting HTML into OfficeIMO.Markdown content.
/// </summary>
public static class HtmlMarkdownConverterExtensions {
    /// <summary>
    /// Converts a shared OfficeIMO HTML conversion document into Markdown text.
    /// </summary>
    /// <param name="document">Shared HTML conversion document.</param>
    /// <param name="options">Optional conversion options. Default options are used when omitted.</param>
    /// <returns>The rendered Markdown text.</returns>
    public static string ToMarkdown(this HtmlConversionDocument document, HtmlToMarkdownOptions? options = null) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        return document.ToMarkdownDocumentResult(operation).Value.ToMarkdown(operation.MarkdownWriteOptions);
    }

    /// <summary>
    /// Converts a shared OfficeIMO HTML conversion document into a Markdown document model.
    /// </summary>
    /// <param name="document">Shared HTML conversion document.</param>
    /// <param name="options">Optional conversion options. Default options are used when omitted.</param>
    /// <returns>A structural <see cref="MarkdownDoc"/> representing the converted Markdown.</returns>
    public static MarkdownDoc ToMarkdownDocument(this HtmlConversionDocument document, HtmlToMarkdownOptions? options = null) {
        return document.ToMarkdownDocumentResult(options).Value;
    }

    /// <summary>Converts a shared HTML conversion document into Markdown with operation-scoped evidence.</summary>
    public static HtmlToMarkdownResult ToMarkdownDocumentResult(
        this HtmlConversionDocument document,
        HtmlToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        operation.BaseUri ??= document.FallbackBaseUri;
        var converter = new HtmlToMarkdownConverter();
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument;
        if (document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint) {
            sourceDocument = document.CreateDocumentForConversion(HtmlCssMediaContext.Print);
        } else {
            sourceDocument = document.CreatePolicyNormalizedDocumentForConversion();
            HtmlActiveMediaFilter.FilterUnsupportedPictureSources(sourceDocument);
        }
        MarkdownDoc value = converter.ConvertToDocument(
            sourceDocument,
            operation);
        return new HtmlToMarkdownResult(value, document.Diagnostics);
    }

    /// <summary>Saves converted Markdown text to a path.</summary>
    public static void SaveAsMarkdown(
        this HtmlConversionDocument document,
        string path,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        document.ToMarkdownDocumentResult(operation).Value.Save(path, operation.MarkdownWriteOptions, encoding);
    }

    /// <summary>Saves converted Markdown text to a caller-owned stream.</summary>
    public static void SaveAsMarkdown(
        this HtmlConversionDocument document,
        Stream stream,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        document.ToMarkdownDocumentResult(operation).Value.Save(stream, operation.MarkdownWriteOptions, encoding);
    }

    /// <summary>Asynchronously saves converted Markdown text to a path.</summary>
    public static Task SaveAsMarkdownAsync(
        this HtmlConversionDocument document,
        string path,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        return document.ToMarkdownDocumentResult(operation).Value
            .SaveAsync(path, operation.MarkdownWriteOptions, encoding, cancellationToken);
    }

    /// <summary>Asynchronously saves converted Markdown text to a caller-owned stream.</summary>
    public static Task SaveAsMarkdownAsync(
        this HtmlConversionDocument document,
        Stream stream,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        return document.ToMarkdownDocumentResult(operation).Value
            .SaveAsync(stream, operation.MarkdownWriteOptions, encoding, cancellationToken);
    }
}
