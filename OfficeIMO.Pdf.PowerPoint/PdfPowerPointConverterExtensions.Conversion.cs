using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.PowerPoint;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.Pdf.PowerPoint;

public static partial class PdfPowerPointConverterExtensions {
    /// <summary>Reconstructs a PowerPoint document from the source PDF.</summary>
    public static PptCore.PowerPointPresentation ToPowerPoint(this PdfDocument document, PowerPointPdfImportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Reconstructs a PowerPoint document and returns its fidelity report.</summary>
    public static PptCore.PowerPointPresentation ToPowerPoint(this PdfDocument document, PowerPointPdfImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfPowerPointConversionResult result = document.ConvertToPowerPointWithEvidence(options, cancellationToken);
        report = PdfConversionEvidence.CreateReport("Pdf", result.Report.Warnings,
            elapsed: timer.Elapsed, itemsWritten: result.Value.Slides.Count, target: "PowerPoint");
        try {
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Asynchronously reconstructs PowerPoint content and returns the model with its report.</summary>
    public static async Task<OfficeConversionResult<PptCore.PowerPointPresentation>> ToPowerPointAsync(this PdfDocument document,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        return await Task.Run(() => {
            PptCore.PowerPointPresentation value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<PptCore.PowerPointPresentation>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reconstructs and saves a PowerPoint file, reporting any fidelity loss.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this PdfDocument document, string path,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using var value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "PowerPoint", Path.GetFullPath(path),
            report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Reconstructs PowerPoint content into a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this PdfDocument document, Stream stream,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using var value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "PowerPoint", itemsWritten: report.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs and writes a PowerPoint file.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this PdfDocument document, string path,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using var value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "PowerPoint", Path.GetFullPath(path),
            report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs PowerPoint content into a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this PdfDocument document, Stream stream,
        PowerPointPdfImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfImportOptions()).Snapshot();
        var timer = Stopwatch.StartNew();
        using var value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "PowerPoint", itemsWritten: report.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }
}
