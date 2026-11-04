using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    // Parse strings through the existing bounded, untrusted HTML source contract.
    private static HtmlConversionDocument ParseWordHtml(string html, HtmlWordExportOptions options, CancellationToken cancellationToken = default) {
        var sourceOptions = new HtmlConversionDocumentOptions {
            IncludeNormalizedHtml = false
        };
        sourceOptions.Limits = HtmlConversionLimits.Intersect(sourceOptions.Limits, options.Limits);
        sourceOptions.UrlPolicy = HtmlUrlPolicy.Intersect(sourceOptions.UrlPolicy, options.HyperlinkUrlPolicy);
        sourceOptions.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(sourceOptions.ResourceUrlPolicy, options.ResourceUrlPolicy);
        return HtmlConversionDocument.Parse(html, sourceOptions, cancellationToken);
    }

    /// <summary>Parses and appends HTML to the document body.</summary>
    /// <param name="doc">Target to modify.</param>
    /// <param name="html">HTML source text, parsed with untrusted input limits and policies.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
    public static IReadOnlyList<WordElement> AddHtml(this WordDocument doc, string html, HtmlWordExportOptions? options = null) {
        ArgumentNullException.ThrowIfNull(doc);
        HtmlWordExportOptions resolved = (options ?? CreateWordOptionsForSharedDocument()).Clone();
        return doc.AddHtml(ParseWordHtml(html, resolved), resolved);
    }

    /// <summary>Parses and asynchronously appends HTML to the document body.</summary>
    /// <param name="doc">Target to modify.</param>
    /// <param name="html">HTML source text, parsed with untrusted input limits and policies.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
    public static Task<IReadOnlyList<WordElement>> AddHtmlAsync(this WordDocument doc, string html, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(doc);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlWordExportOptions resolved = (options ?? CreateWordOptionsForSharedDocument()).Clone();
        return doc.AddHtmlAsync(ParseWordHtml(html, resolved, cancellationToken), resolved, cancellationToken);
    }

    /// <summary>Parses and appends HTML to the table cell.</summary>
    /// <param name="cell">Target to modify.</param>
    /// <param name="html">HTML source text, parsed with untrusted input limits and policies.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
    public static IReadOnlyList<WordElement> AddHtml(this WordTableCell cell, string html, HtmlWordExportOptions? options = null) {
        ArgumentNullException.ThrowIfNull(cell);
        HtmlWordExportOptions resolved = (options ?? CreateWordOptionsForSharedDocument()).Clone();
        return cell.AddHtml(ParseWordHtml(html, resolved), resolved);
    }

    /// <summary>Parses and asynchronously appends HTML to the table cell.</summary>
    /// <param name="cell">Target to modify.</param>
    /// <param name="html">HTML source text, parsed with untrusted input limits and policies.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
    public static Task<IReadOnlyList<WordElement>> AddHtmlAsync(this WordTableCell cell, string html, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(cell);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlWordExportOptions resolved = (options ?? CreateWordOptionsForSharedDocument()).Clone();
        return cell.AddHtmlAsync(ParseWordHtml(html, resolved, cancellationToken), resolved, cancellationToken);
    }

}
