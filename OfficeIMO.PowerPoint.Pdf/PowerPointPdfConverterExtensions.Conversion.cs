using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    /// <summary>Converts the source document to a PDF model.</summary>
    public static PdfDocument ToPdf(this PptCore.PowerPointPresentation presentation, PowerPointPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        presentation.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the source document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this PptCore.PowerPointPresentation presentation, PowerPointPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = presentation.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, "PowerPoint", options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the source document, returning the model and report together.</summary>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this PptCore.PowerPointPresentation presentation,
        PowerPointPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PowerPointPdfExportOptions()).Clone();
        return await Task.Run(() => {
            PdfDocument value = presentation.ToPdf(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this PptCore.PowerPointPresentation presentation, string path,
        PowerPointPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new PowerPointPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = presentation.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "PowerPoint", path, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "PowerPoint", () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport("PowerPoint", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this PptCore.PowerPointPresentation presentation, Stream stream,
        PowerPointPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = presentation.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "PowerPoint", stream, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "PowerPoint", () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport("PowerPoint", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes a PDF file and returns the complete conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this PptCore.PowerPointPresentation presentation, string path,
        PowerPointPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new PowerPointPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = presentation.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "PowerPoint", path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "PowerPoint", () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("PowerPoint", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes PDF to a caller-owned stream and returns its conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this PptCore.PowerPointPresentation presentation, Stream stream,
        PowerPointPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new PowerPointPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = presentation.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "PowerPoint", stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "PowerPoint", () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("PowerPoint", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
