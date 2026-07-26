using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    /// <summary>
    /// Appends HTML content directly to a Word table cell.
    /// </summary>
    /// <param name="cell">Table cell to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    public static void AddHtml(
        this WordTableCell cell,
        HtmlConversionDocument htmlDocument,
        HtmlToWordOptions? options = null) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        EnsureOfflineSynchronousImport(htmlDocument, resolved);
        cell.AddHtmlAsync(htmlDocument, resolved).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Asynchronously appends HTML content directly to a Word table cell.
    /// </summary>
    /// <param name="cell">Table cell to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    public static async Task AddHtmlAsync(
        this WordTableCell cell,
        HtmlConversionDocument htmlDocument,
        HtmlToWordOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        cancellationToken.ThrowIfCancellationRequested();

        HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        var converter = new HtmlToWordConverter();
        await converter.AddHtmlToTableCellAsync(
            cell,
            CreateWordSourceDocument(htmlDocument, resolved.ConversionReport),
            resolved,
            cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }
}
