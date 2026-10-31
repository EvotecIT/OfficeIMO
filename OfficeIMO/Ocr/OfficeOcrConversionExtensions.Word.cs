using System.Diagnostics;
using OfficeIMO.Word;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO;

/// <summary>Converts accepted recognition content through the native reconstruction owners.</summary>
public static partial class OfficeOcrConversionExtensions {
    /// <summary>Reconstructs Word content from accepted OCR words and detected tables.</summary>
    public static WordDocument ToWord(this PdfOcrMergeResult source, WordPdfImportOptions? options = null,
        CancellationToken cancellationToken = default) => source.ToWord(options, out _, cancellationToken);

    /// <summary>Reconstructs Word content with its completed fidelity report.</summary>
    public static WordDocument ToWord(this PdfOcrMergeResult source, WordPdfImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        source.RequireAcceptedOcrContent();
        var timer = Stopwatch.StartNew();
        var converted = source.Document.ConvertToWordWithEvidence(options, cancellationToken);
        WordDocument value = converted.Value;
        try {
            report = PdfConversionEvidence.CreateReport("Pdf", new IOfficeConversionReport[] { converted.Report },
                elapsed: timer.Elapsed, itemsWritten: value.Sections.Count, target: "Word");
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return value;
        } catch { value.Dispose(); throw; }
    }

    /// <summary>Asynchronously reconstructs Word content and returns the model with its report.</summary>
    public static async Task<OfficeConversionResult<WordDocument>> ToWordAsync(this PdfOcrMergeResult source,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        return await Task.Run(() => {
            WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<WordDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves accepted OCR content as Word, retaining reconstruction findings.</summary>
    public static OfficeConversionReport SaveAsWord(this PdfOcrMergeResult source, string path,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "Word", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes accepted OCR content as Word to a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsWord(this PdfOcrMergeResult source, Stream stream,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "Word", itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously saves accepted OCR content as Word.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this PdfOcrMergeResult source, string path,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Word", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously writes accepted OCR content as Word to a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this PdfOcrMergeResult source, Stream stream,
        WordPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfImportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Word", itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }
}
