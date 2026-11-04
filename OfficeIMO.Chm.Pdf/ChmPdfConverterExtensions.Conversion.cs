using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.Chm;

public static partial class ChmPdfConverterExtensions {
    private const string Source = "Chm";

    /// <summary>Converts the compiled help document to a PDF model.</summary>
    public static PdfDocument ToPdf(this ChmDocument document, ChmPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        document.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the compiled help document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this ChmDocument document, ChmPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, Source, options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the compiled help document, returning the model and report together.</summary>
    /// <remarks>The source pipeline runs asynchronously; PDF serialization completes on the continuation.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this ChmDocument document,
        ChmPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new ChmPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await document.ConvertToPdfWithEvidenceAsync(options, cancellationToken).ConfigureAwait(false);
        PdfDocument value = PdfConversionEvidence.CompleteModel(converted, Source, options.FailOnLoss,
            timer, cancellationToken, out OfficeConversionReport report);
        return new OfficeConversionResult<PdfDocument>(value, report);
    }

    /// <summary>Writes the compiled help document as a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this ChmDocument document, string path,
        ChmPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, path, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes the compiled help document as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this ChmDocument document, Stream stream,
        ChmPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, stream, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the compiled help document as a PDF file and returns the complete conversion report.</summary>
    /// <remarks>The source pipeline runs asynchronously; PDF serialization completes on the continuation. The file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this ChmDocument document, string path,
        ChmPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new ChmPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await document.ConvertToPdfWithEvidenceAsync(options, cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the compiled help document as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    /// <remarks>The source pipeline runs asynchronously; PDF serialization completes on the continuation. The stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this ChmDocument document, Stream stream,
        ChmPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new ChmPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await document.ConvertToPdfWithEvidenceAsync(options, cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
