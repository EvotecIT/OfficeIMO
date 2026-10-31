using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts the source document to a PDF model.</summary>
    public static PdfDocument ToPdf(this RtfDocument document, RtfToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        document.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the source document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this RtfDocument document, RtfToPdfOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, "Rtf", options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the source document, returning the model and report together.</summary>
    public static Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this RtfDocument document,
        RtfToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => {
            PdfDocument value = document.ToPdf(options, out OfficeConversionReport report, cancellationToken);
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken);

    /// <summary>Writes a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this RtfDocument document, string path,
        RtfToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "Rtf", path, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "Rtf", () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport("Rtf", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this RtfDocument document, Stream stream,
        RtfToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "Rtf", stream, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "Rtf", () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport("Rtf", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes a PDF file and returns the complete conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this RtfDocument document, string path,
        RtfToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Rtf", path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "Rtf", () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Rtf", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes PDF to a caller-owned stream and returns its conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this RtfDocument document, Stream stream,
        RtfToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Rtf", stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "Rtf", () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Rtf", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
