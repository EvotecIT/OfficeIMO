using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

public static partial class MarkdownPdfConverterExtensions {
    /// <summary>Converts the source document to a PDF model.</summary>
    public static PdfDocument ToPdf(this MarkdownDocument document, MarkdownPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        document.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the source document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this MarkdownDocument document, MarkdownPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, "Markdown", options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the source document, returning the model and report together.</summary>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this MarkdownDocument document,
        MarkdownPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new MarkdownPdfExportOptions()).Clone();
        return await Task.Run(() => {
            PdfDocument value = document.ToPdf(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this MarkdownDocument document, string path,
        MarkdownPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new MarkdownPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "Markdown", path, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "Markdown", () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport("Markdown", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this MarkdownDocument document, Stream stream,
        MarkdownPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "Markdown", stream, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "Markdown", () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport("Markdown", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes a PDF file and returns the complete conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this MarkdownDocument document, string path,
        MarkdownPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new MarkdownPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Markdown", path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "Markdown", () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Markdown", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes PDF to a caller-owned stream and returns its conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this MarkdownDocument document, Stream stream,
        MarkdownPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new MarkdownPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Markdown", stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "Markdown", () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Markdown", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
