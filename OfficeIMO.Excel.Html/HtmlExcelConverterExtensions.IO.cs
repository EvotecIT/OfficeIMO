using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    /// <summary>Reads semantic HTML from a stream and imports a workbook.</summary>
    public static ExcelDocument ToExcelDocument(this Stream htmlStream, HtmlToExcelOptions? options = null) =>
        HtmlTextIO.Read(htmlStream).ToExcelDocument(options);

    /// <summary>Reads semantic HTML from a stream and returns a workbook plus structured evidence.</summary>
    public static HtmlToExcelResult ToExcelDocumentResult(this Stream htmlStream, HtmlToExcelOptions? options = null) =>
        HtmlTextIO.Read(htmlStream).ToExcelDocumentResult(options);

    /// <summary>Asynchronously reads semantic HTML from a stream and imports a workbook.</summary>
    public static async Task<ExcelDocument> ToExcelDocumentAsync(this Stream htmlStream, HtmlToExcelOptions? options = null, CancellationToken cancellationToken = default) =>
        (await HtmlTextIO.ReadAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToExcelDocument(options);

    /// <summary>Asynchronously reads semantic HTML from a stream and returns structured evidence.</summary>
    public static async Task<HtmlToExcelResult> ToExcelDocumentResultAsync(this Stream htmlStream, HtmlToExcelOptions? options = null, CancellationToken cancellationToken = default) =>
        (await HtmlTextIO.ReadAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToExcelDocumentResult(options);
}
