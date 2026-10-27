using System.Diagnostics;
using OfficeIMO.PowerPoint;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO;

/// <summary>Converts accepted recognition content through the native reconstruction owners.</summary>
public static partial class OfficeOcrConversionExtensions {
    /// <summary>Reconstructs PowerPoint content from accepted OCR words and detected tables.</summary>
    public static PowerPointPresentation ToPowerPoint(this PdfOcrMergeResult source, PowerPointPdfImportOptions? options = null,
        CancellationToken cancellationToken = default) => source.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Reconstructs PowerPoint content with its completed fidelity report.</summary>
    public static PowerPointPresentation ToPowerPoint(this PdfOcrMergeResult source, PowerPointPdfImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        source.RequireAcceptedOcrContent();
        var timer = Stopwatch.StartNew();
        var converted = source.Document.ConvertToPowerPointWithEvidence(options, cancellationToken);
        PowerPointPresentation value = converted.Value;
        try {
            report = PdfConversionEvidence.CreateReport("Pdf", new IOfficeConversionReport[] { converted.Report },
                elapsed: timer.Elapsed, itemsWritten: value.Slides.Count, target: "PowerPoint");
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return value;
        } catch { value.Dispose(); throw; }
    }

    /// <summary>Asynchronously reconstructs PowerPoint content and returns the model with its report.</summary>
    public static async Task<OfficeConversionResult<PowerPointPresentation>> ToPowerPointAsync(this PdfOcrMergeResult source,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        return await Task.Run(() => {
            PowerPointPresentation value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<PowerPointPresentation>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves accepted OCR content as PowerPoint, retaining reconstruction findings.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this PdfOcrMergeResult source, string path,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using PowerPointPresentation value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "PowerPoint", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes accepted OCR content as PowerPoint to a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this PdfOcrMergeResult source, Stream stream,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using PowerPointPresentation value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "PowerPoint", itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously saves accepted OCR content as PowerPoint.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this PdfOcrMergeResult source, string path,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using PowerPointPresentation value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "PowerPoint", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously writes accepted OCR content as PowerPoint to a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this PdfOcrMergeResult source, Stream stream,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using PowerPointPresentation value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "PowerPoint", itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }
}
