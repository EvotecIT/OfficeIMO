using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    /// <summary>Converts a Numbers spreadsheet and returns the Excel model together with the typed source projection and loss report.</summary>
    /// <remarks>The result keeps the opened source package and its typed projection, which the common report does not carry.
    /// Use <see cref="NumbersToExcelResult.RequireCompleteEditableReconstruction"/> to reject visual fallbacks and incomplete reconstruction.</remarks>
    public static NumbersToExcelResult ToExcelResult(this IWorkSourceDocument package, IWorkConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (package == null) throw new ArgumentNullException(nameof(package));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectNumbers(package, options);
    }

    /// <summary>Converts a Numbers spreadsheet into the editable Excel model.</summary>
    public static ExcelDocument ToExcel(this IWorkSourceDocument package, NumbersExcelExportOptions? options = null, CancellationToken cancellationToken = default) =>
        package.ToExcel(options, out _, cancellationToken);

    /// <summary>Converts a Numbers spreadsheet into the editable Excel model and returns the completed conversion report.</summary>
    public static ExcelDocument ToExcel(this IWorkSourceDocument package, NumbersExcelExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (package == null) throw new ArgumentNullException(nameof(package));
        cancellationToken.ThrowIfCancellationRequested();
        NumbersToExcelResult result = package.ToExcelResult(options?.ConversionOptions, cancellationToken);
        ExcelDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a Numbers spreadsheet into the editable Excel model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this IWorkSourceDocument package, NumbersExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (package.ToExcel(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this IWorkSourceDocument package, string path, NumbersExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = package.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this IWorkSourceDocument package, Stream stream, NumbersExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = package.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this IWorkSourceDocument package, string path, NumbersExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using ExcelDocument value = package.ToExcel(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this IWorkSourceDocument package, Stream stream, NumbersExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using ExcelDocument value = package.ToExcel(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
