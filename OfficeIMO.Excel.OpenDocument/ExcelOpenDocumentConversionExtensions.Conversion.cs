using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenDocument;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    /// <summary>Converts a workbook to an in-memory ODS spreadsheet; the LossPolicy option rejects loss.</summary>
    public static OdsDocument ToOpenDocument(this ExcelDocument source, ExcelOpenDocumentConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        source.ToOpenDocument(options, out _, cancellationToken);

    /// <summary>Converts a workbook to an in-memory ODS spreadsheet; the LossPolicy option rejects loss and returns the completed conversion report.</summary>
    public static OdsDocument ToOpenDocument(this ExcelDocument source, ExcelOpenDocumentConversionOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        OdfConversionResult<OdsDocument> result = source.ConvertToOpenDocumentWithEvidence(options);
        OdsDocument value = result.Value;
        OfficeConversionReport evidence = result.Report;
        return OfficeConversionOutput.Complete(value, evidence, false, out report);
    }

    /// <summary>Converts a workbook to an in-memory ODS spreadsheet; the LossPolicy option rejects loss, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<OdsDocument>> ToOpenDocumentAsync(this ExcelDocument source, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsOpenDocument(this ExcelDocument source, string path, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using OdsDocument value = source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsOpenDocument(this ExcelDocument source, Stream stream, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using OdsDocument value = source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOpenDocumentAsync(this ExcelDocument source, string path, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using OdsDocument value = source.ToOpenDocument(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOpenDocumentAsync(this ExcelDocument source, Stream stream, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using OdsDocument value = source.ToOpenDocument(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts an ODS spreadsheet to an in-memory workbook; the LossPolicy option rejects loss.</summary>
    public static ExcelDocument ToExcel(this OdsDocument source, ExcelOpenDocumentConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        source.ToExcel(options, out _, cancellationToken);

    /// <summary>Converts an ODS spreadsheet to an in-memory workbook; the LossPolicy option rejects loss and returns the completed conversion report.</summary>
    public static ExcelDocument ToExcel(this OdsDocument source, ExcelOpenDocumentConversionOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        OdfConversionResult<ExcelDocument> result = source.ConvertToExcelWithEvidence(options);
        ExcelDocument value = result.Value;
        OfficeConversionReport evidence = result.Report;
        return OfficeConversionOutput.Complete(value, evidence, false, out report);
    }

    /// <summary>Converts an ODS spreadsheet to an in-memory workbook; the LossPolicy option rejects loss, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this OdsDocument source, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (source.ToExcel(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
