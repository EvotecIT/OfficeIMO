using OfficeIMO.Markdown;
using OfficeIMO.Html;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Extension helpers for converting HTML into OfficeIMO.Markdown content.
/// </summary>
public static partial class HtmlMarkdownConverterExtensions {












    internal static HtmlToMarkdownResult ConvertToMarkdownWithEvidence(
        this HtmlConversionDocument document,
        MarkdownHtmlImportOptions operation, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ApplyDocumentPolicies(document, operation);
        var converter = new HtmlToMarkdownConverter(cancellationToken);
        MarkdownDocument value;
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
        AngleSharp.Html.Dom.IHtmlDocument prepared, MarkdownHtmlImportOptions options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        MarkdownHtmlImportOptions operation = options.Clone();
        ApplyDocumentPolicies(document, operation, documentTransformsApplied: true);
        if (document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint) {
            HtmlActiveMediaFilter.Filter(prepared, HtmlCssMediaContext.Print, diagnostics: null, cancellationToken);
        } else {
            HtmlActiveMediaFilter.FilterUnsupportedPictureSources(prepared, cancellationToken);
        }
        return new HtmlToMarkdownConverter(cancellationToken).ConvertReadOnlyDocumentToDocument(
            prepared, operation, document.SourceHtml.Length).ToMarkdown(operation.MarkdownWriteOptions, cancellationToken);
    }

    private static void ApplyDocumentPolicies(HtmlConversionDocument document, MarkdownHtmlImportOptions operation,
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
        MarkdownHtmlImportOptions operation) =>
        document.ProfileContract.Profile != HtmlConversionProfile.HighFidelityPrint
        && operation.ExcludeSelectors.Count == 0
        && operation.ElementFilters.Count == 0
        && operation.ElementBlockConverters.Count == 0
        && operation.InlineElementConverters.Count == 0
        && document.SourceHtml.IndexOf("<picture", StringComparison.OrdinalIgnoreCase) < 0;








}
