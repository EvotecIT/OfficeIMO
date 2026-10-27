using System.Diagnostics;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO;

/// <summary>Converts accepted recognition content through the native reconstruction owners.</summary>
public static partial class OfficeOcrConversionExtensions {
    /// <summary>Reconstructs Excel content from accepted OCR words and detected tables.</summary>
    public static ExcelDocument ToExcel(this PdfOcrMergeResult source, ExcelPdfImportOptions? options = null,
        CancellationToken cancellationToken = default) => source.ToExcel(options, out _, cancellationToken);

    /// <summary>Reconstructs Excel content with its completed fidelity report.</summary>
    public static ExcelDocument ToExcel(this PdfOcrMergeResult source, ExcelPdfImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        source.RequireAcceptedOcrContent();
        var timer = Stopwatch.StartNew();
        var converted = source.Document.ConvertToExcelWithEvidence(options, cancellationToken);
        ExcelDocument value = converted.Value;
        try {
            report = PdfConversionEvidence.CreateReport("Pdf", new IOfficeConversionReport[] { converted.Report },
                elapsed: timer.Elapsed, itemsWritten: value.Sheets.Count, target: "Excel");
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return value;
        } catch { value.Dispose(); throw; }
    }

    /// <summary>Asynchronously reconstructs Excel content and returns the model with its report.</summary>
    public static async Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this PdfOcrMergeResult source,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        return await Task.Run(() => {
            ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<ExcelDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves accepted OCR content as Excel, retaining reconstruction findings.</summary>
    public static OfficeConversionReport SaveAsExcel(this PdfOcrMergeResult source, string path,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "Excel", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes accepted OCR content as Excel to a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsExcel(this PdfOcrMergeResult source, Stream stream,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "Excel", itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously saves accepted OCR content as Excel.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this PdfOcrMergeResult source, string path,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Excel", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously writes accepted OCR content as Excel to a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this PdfOcrMergeResult source, Stream stream,
        ExcelPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new ExcelPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using ExcelDocument value = source.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Excel", itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }
}
