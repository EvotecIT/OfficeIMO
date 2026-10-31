using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Pdf.Excel;

public static partial class PdfExcelTableConverterExtensions {
    /// <summary>Reconstructs a Excel document from the source PDF.</summary>
    public static ExcelDocument ToExcel(this PdfDocument document, ExcelPdfImportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToExcel(options, out _, cancellationToken);

    /// <summary>Reconstructs a Excel document and returns its fidelity report.</summary>
    public static ExcelDocument ToExcel(this PdfDocument document, ExcelPdfImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfExcelTableImportResult result = document.ConvertToExcelWithEvidence(options, cancellationToken);
        report = PdfConversionEvidence.CreateReport("Pdf", new[] { result.Report },
            elapsed: timer.Elapsed, itemsWritten: result.Value.Sheets.Count, target: "Excel");
        try {
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Asynchronously reconstructs Excel content and returns the model with its report.</summary>
    public static async Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this PdfDocument document,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        return await Task.Run(() => {
            ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<ExcelDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reconstructs and saves a Excel file, reporting any fidelity loss.</summary>
    public static OfficeConversionReport SaveAsExcel(this PdfDocument document, string path,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using var value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "Excel", Path.GetFullPath(path),
            report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Reconstructs Excel content into a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsExcel(this PdfDocument document, Stream stream,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using var value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "Excel", itemsWritten: report.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs and writes a Excel file.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this PdfDocument document, string path,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using var value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Excel", Path.GetFullPath(path),
            report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs Excel content into a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this PdfDocument document, Stream stream,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using var value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Excel", itemsWritten: report.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }
}
