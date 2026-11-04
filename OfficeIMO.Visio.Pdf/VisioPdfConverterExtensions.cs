using System.Diagnostics;
using OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

/// <summary>Converts loaded Visio documents through the Visio-owned neutral projection and PDF engine.</summary>
public static class VisioPdfConverterExtensions {
    private const string Source = "Visio";

    /// <summary>Converts a Visio document to a PDF model.</summary>
    public static PdfDocument ToPdf(this VisioDocument document, VisioPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        document.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts a Visio document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this VisioDocument document, VisioPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, Source, options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts a Visio document, returning the model and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this VisioDocument document,
        VisioPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new VisioPdfExportOptions()).Clone();
        return await Task.Run(() => {
            PdfDocument value = document.ToPdf(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this VisioDocument document, string path,
        VisioPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, path, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this VisioDocument document, Stream stream,
        VisioPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, stream, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes a PDF file and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this VisioDocument document, string path,
        VisioPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new VisioPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => document.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this VisioDocument document, Stream stream,
        VisioPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new VisioPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => document.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    internal static PdfDocumentConversionResult ConvertToPdfWithEvidence(this VisioDocument document,
        VisioPdfExportOptions? options, CancellationToken cancellationToken) =>
        VisioPdfConversionEngine.Convert(document, options, cancellationToken);
}
