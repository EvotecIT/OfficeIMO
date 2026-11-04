using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    /// <summary>Converts prepared HTML into a native workbook.</summary>
    public static ExcelDocument ToExcel(this HtmlConversionDocument document, HtmlExcelExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToExcel(options, out _, cancellationToken);

    /// <summary>Converts prepared HTML into a native workbook and returns the completed conversion report.</summary>
    public static ExcelDocument ToExcel(this HtmlConversionDocument document, HtmlExcelExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToExcelResult result = document.ConvertToExcelWithEvidence(options);
        ExcelDocument value = GetWorkbookOrThrow(result);
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("Html", "Excel", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts prepared HTML into a native workbook, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this HtmlConversionDocument document, HtmlExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToExcel(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this HtmlConversionDocument document, string path, HtmlExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this HtmlConversionDocument document, Stream stream, HtmlExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this HtmlConversionDocument document, string path, HtmlExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this HtmlConversionDocument document, Stream stream, HtmlExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
