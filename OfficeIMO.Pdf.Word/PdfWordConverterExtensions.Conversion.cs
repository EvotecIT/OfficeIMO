using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Pdf.Word;

public static partial class PdfWordConverterExtensions {
    /// <summary>Reconstructs a Word document from the source PDF.</summary>
    public static WordDocument ToWord(this PdfDocument document, WordPdfImportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToWord(options, out _, cancellationToken);

    /// <summary>Reconstructs a Word document and returns its fidelity report.</summary>
    public static WordDocument ToWord(this PdfDocument document, WordPdfImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfWordConversionResult result = document.ConvertToWordWithEvidence(options, cancellationToken);
        report = PdfConversionEvidence.CreateReport("Pdf", result.Report.Warnings,
            elapsed: timer.Elapsed, itemsWritten: result.Value.Sections.Count, target: "Word");
        try {
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Asynchronously reconstructs Word content and returns the model with its report.</summary>
    public static async Task<OfficeConversionResult<WordDocument>> ToWordAsync(this PdfDocument document,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        return await Task.Run(() => {
            WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<WordDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reconstructs and saves a Word file, reporting any fidelity loss.</summary>
    public static OfficeConversionReport SaveAsWord(this PdfDocument document, string path,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using var value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "Word", Path.GetFullPath(path),
            report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Reconstructs Word content into a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsWord(this PdfDocument document, Stream stream,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using var value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "Word", itemsWritten: report.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs and writes a Word file.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this PdfDocument document, string path,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using var value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Word", Path.GetFullPath(path),
            report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs Word content into a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this PdfDocument document, Stream stream,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using var value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Word", itemsWritten: report.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }
}
