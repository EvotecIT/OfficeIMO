using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    /// <summary>Writes a workbook as UTF-8 HTML to a stream and leaves the stream open.</summary>
    public static void SaveAsHtml(this ExcelDocument workbook, Stream stream, ExcelHtmlSaveOptions? options = null) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        HtmlTextIO.Write(stream, workbook.ToHtml(options));
    }

    /// <summary>Writes a worksheet as UTF-8 HTML to a stream and leaves the stream open.</summary>
    public static void SaveAsHtml(this ExcelSheet sheet, Stream stream, ExcelHtmlSaveOptions? options = null) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        HtmlTextIO.Write(stream, sheet.ToHtml(options));
    }

    /// <summary>Asynchronously writes a workbook as UTF-8 HTML.</summary>
    public static async Task SaveAsHtmlAsync(this ExcelDocument workbook, string path, ExcelHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        cancellationToken.ThrowIfCancellationRequested();
        await HtmlTextIO.WriteAsync(path, workbook.ToHtml(options), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes a workbook as UTF-8 HTML to a stream and leaves it open.</summary>
    public static async Task SaveAsHtmlAsync(this ExcelDocument workbook, Stream stream, ExcelHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        cancellationToken.ThrowIfCancellationRequested();
        await HtmlTextIO.WriteAsync(stream, workbook.ToHtml(options), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes a worksheet as UTF-8 HTML.</summary>
    public static async Task SaveAsHtmlAsync(this ExcelSheet sheet, string path, ExcelHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        cancellationToken.ThrowIfCancellationRequested();
        await HtmlTextIO.WriteAsync(path, sheet.ToHtml(options), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes a worksheet as UTF-8 HTML to a stream and leaves it open.</summary>
    public static async Task SaveAsHtmlAsync(this ExcelSheet sheet, Stream stream, ExcelHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        cancellationToken.ThrowIfCancellationRequested();
        await HtmlTextIO.WriteAsync(stream, sheet.ToHtml(options), cancellationToken).ConfigureAwait(false);
    }
}
