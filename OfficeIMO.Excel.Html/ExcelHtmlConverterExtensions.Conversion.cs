using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    /// <summary>Converts a workbook to HTML text.</summary>
    public static string ToHtml(this ExcelDocument workbook, ExcelHtmlExportOptions? options = null, CancellationToken cancellationToken = default) =>
        workbook.ToHtml(options, out _, cancellationToken);

    /// <summary>Converts a workbook to HTML text and returns the completed conversion report.</summary>
    public static string ToHtml(this ExcelDocument workbook, ExcelHtmlExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlTextConversionResult result = workbook.ConvertToHtmlWithEvidence(options);
        string value = result.RequireValue();
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("Excel", "Html", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a workbook to HTML text, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<string>> ToHtmlAsync(this ExcelDocument workbook, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (workbook.ToHtml(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this ExcelDocument workbook, string path, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = workbook.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this ExcelDocument workbook, Stream stream, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = workbook.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this ExcelDocument workbook, string path, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = workbook.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this ExcelDocument workbook, Stream stream, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = workbook.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts a worksheet to HTML text.</summary>
    public static string ToHtml(this ExcelSheet sheet, ExcelHtmlExportOptions? options = null, CancellationToken cancellationToken = default) =>
        sheet.ToHtml(options, out _, cancellationToken);

    /// <summary>Converts a worksheet to HTML text and returns the completed conversion report.</summary>
    public static string ToHtml(this ExcelSheet sheet, ExcelHtmlExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlTextConversionResult result = sheet.ConvertToHtmlWithEvidence(options);
        string value = result.RequireValue();
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("Excel", "Html", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a worksheet to HTML text, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<string>> ToHtmlAsync(this ExcelSheet sheet, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (sheet.ToHtml(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this ExcelSheet sheet, string path, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = sheet.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this ExcelSheet sheet, Stream stream, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = sheet.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this ExcelSheet sheet, string path, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = sheet.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this ExcelSheet sheet, Stream stream, ExcelHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = sheet.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

}
