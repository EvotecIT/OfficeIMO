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
    public static string ToMarkdown(this HtmlConversionDocument document, HtmlToMarkdownOptions? options = null) =>
        document.ToMarkdown(options, CancellationToken.None);

    /// <summary>Converts shared HTML to Markdown with cooperative cancellation during projection and serialization.</summary>
    public static string ToMarkdown(this HtmlConversionDocument document, HtmlToMarkdownOptions? options, CancellationToken cancellationToken) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        return ToMarkdownDocumentResultCore(document, operation, cancellationToken).Value.ToMarkdown(operation.MarkdownWriteOptions, cancellationToken);
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

    /// <summary>Converts shared HTML to a Markdown model with cooperative cancellation during projection.</summary>
    public static MarkdownDoc ToMarkdownDocument(this HtmlConversionDocument document, HtmlToMarkdownOptions? options, CancellationToken cancellationToken) =>
        document.ToMarkdownDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a shared HTML conversion document into Markdown with operation-scoped evidence.</summary>
    public static HtmlToMarkdownResult ToMarkdownDocumentResult(
        this HtmlConversionDocument document,
        HtmlToMarkdownOptions? options = null) => document.ToMarkdownDocumentResult(options, CancellationToken.None);

    /// <summary>Converts shared HTML into Markdown with operation-scoped evidence and cooperative cancellation.</summary>
    public static HtmlToMarkdownResult ToMarkdownDocumentResult(this HtmlConversionDocument document,
        HtmlToMarkdownOptions? options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        return ToMarkdownDocumentResultCore(document, operation, cancellationToken);
    }

    private static HtmlToMarkdownResult ToMarkdownDocumentResultCore(
        HtmlConversionDocument document,
        HtmlToMarkdownOptions operation, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ApplyDocumentPolicies(document, operation);
        var converter = new HtmlToMarkdownConverter(cancellationToken);
        MarkdownDoc value;
        if (CanProjectSourceReadOnly(document, operation)) {
            value = document.ProjectSourceDocument(sourceDocument =>
                converter.ConvertReadOnlyDocumentToDocument(
                    sourceDocument,
                    operation,
                    document.SourceHtml.Length), cancellationToken);
        } else {
            AngleSharp.Html.Dom.IHtmlDocument sourceDocument = document.CreateSourceDocumentForConversion(cancellationToken);
            if (document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint) {
                HtmlActiveMediaFilter.Filter(sourceDocument, HtmlCssMediaContext.Print, diagnostics: null, cancellationToken);
            } else {
                HtmlActiveMediaFilter.FilterUnsupportedPictureSources(sourceDocument, cancellationToken);
            }
            value = converter.ConvertPreparedDocumentToDocument(
                sourceDocument,
                operation,
                document.SourceHtml.Length);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlToMarkdownResult(value, document.Diagnostics.Concat(converter.Diagnostics));
    }

    /// <summary>Projects an independently owned, already filtered DOM using the source document's URL policies.</summary>
    internal static string ToMarkdownPreparedDocument(HtmlConversionDocument document,
        AngleSharp.Html.Dom.IHtmlDocument prepared, HtmlToMarkdownOptions options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToMarkdownOptions operation = options.Clone();
        ApplyDocumentPolicies(document, operation, documentTransformsApplied: true);
        if (document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint) {
            HtmlActiveMediaFilter.Filter(prepared, HtmlCssMediaContext.Print, diagnostics: null, cancellationToken);
        } else {
            HtmlActiveMediaFilter.FilterUnsupportedPictureSources(prepared, cancellationToken);
        }
        return new HtmlToMarkdownConverter(cancellationToken).ConvertReadOnlyDocumentToDocument(
            prepared, operation, document.SourceHtml.Length).ToMarkdown(operation.MarkdownWriteOptions, cancellationToken);
    }

    private static void ApplyDocumentPolicies(HtmlConversionDocument document, HtmlToMarkdownOptions operation,
        bool documentTransformsApplied = false) {
        operation.BaseUri ??= document.FallbackBaseUri;
        HtmlUrlPolicy requestedHyperlinkPolicy = operation.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        HtmlUrlPolicy requestedResourcePolicy = operation.ResourceUrlPolicy ?? requestedHyperlinkPolicy;
        HtmlUrlPolicy documentHyperlinkPolicy = document.HyperlinkUrlPolicy;
        HtmlUrlPolicy documentResourcePolicy = document.ResourceUrlPolicy;
        if (documentTransformsApplied) {
            // Prepared DOM URLs have already passed the document transforms. Retain
            // their restrictions while applying only the requested projection transforms.
            documentHyperlinkPolicy = documentHyperlinkPolicy.Clone();
            documentResourcePolicy = documentResourcePolicy.Clone();
            documentHyperlinkPolicy.ResolvedUrlTransform = null;
            documentResourcePolicy.ResolvedUrlTransform = null;
        }
        operation.UrlPolicy = HtmlUrlPolicy.Intersect(documentHyperlinkPolicy, requestedHyperlinkPolicy);
        operation.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(documentResourcePolicy, requestedResourcePolicy);
    }

    private static bool CanProjectSourceReadOnly(
        HtmlConversionDocument document,
        HtmlToMarkdownOptions operation) =>
        document.ProfileContract.Profile != HtmlConversionProfile.HighFidelityPrint
        && operation.ExcludeSelectors.Count == 0
        && operation.ElementFilters.Count == 0
        && operation.ElementBlockConverters.Count == 0
        && operation.InlineElementConverters.Count == 0
        && document.SourceHtml.IndexOf("<picture", StringComparison.OrdinalIgnoreCase) < 0;

    /// <summary>Saves converted Markdown text to a path.</summary>
    public static void SaveAsMarkdown(
        this HtmlConversionDocument document,
        string path,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        ToMarkdownDocumentResultCore(document, operation).Value.Save(path, operation.MarkdownWriteOptions, encoding);
    }

    /// <summary>Saves converted Markdown text to a caller-owned stream.</summary>
    public static void SaveAsMarkdown(
        this HtmlConversionDocument document,
        Stream stream,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        ToMarkdownDocumentResultCore(document, operation).Value.Save(stream, operation.MarkdownWriteOptions, encoding);
    }

    /// <summary>Asynchronously saves converted Markdown text to a path.</summary>
    public static Task SaveAsMarkdownAsync(
        this HtmlConversionDocument document,
        string path,
        HtmlToMarkdownOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        HtmlToMarkdownOptions operation = options?.Clone() ?? new HtmlToMarkdownOptions();
        return ToMarkdownDocumentResultCore(document, operation, cancellationToken).Value
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
        return ToMarkdownDocumentResultCore(document, operation, cancellationToken).Value
            .SaveAsync(stream, operation.MarkdownWriteOptions, encoding, cancellationToken);
    }
}
