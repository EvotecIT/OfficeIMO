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
    /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
    public static IReadOnlyList<WordElement> AddHtml(
        this WordTableCell cell,
        HtmlConversionDocument htmlDocument,
        WordHtmlImportOptions? options = null) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        EnsureOfflineSynchronousImport(htmlDocument, resolved);
        return cell.AddHtmlAsync(htmlDocument, resolved).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Asynchronously appends HTML content directly to a Word table cell.
    /// </summary>
    /// <param name="cell">Table cell to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
    public static async Task<IReadOnlyList<WordElement>> AddHtmlAsync(
        this WordTableCell cell,
        HtmlConversionDocument htmlDocument,
        WordHtmlImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        cancellationToken.ThrowIfCancellationRequested();

        WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        var converter = new HtmlToWordConverter();
        IReadOnlyList<WordElement> elements = await converter.AddHtmlToTableCellAsync(
            cell,
            CreateWordSourceDocument(htmlDocument, resolved.ConversionReport),
            resolved,
            cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return elements;
    }
}
