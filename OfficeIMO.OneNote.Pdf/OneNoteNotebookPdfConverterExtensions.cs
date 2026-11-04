using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Converts offline OneNote notebooks to semantic PDF documents.</summary>
public static class OneNoteNotebookPdfConverterExtensions {
    private const string Source = "OneNote";

    internal static PdfDocumentConversionResult ConvertToPdfWithEvidence(this OneNoteNotebook notebook, OneNotePdfExportOptions? options,
        CancellationToken cancellationToken) =>
        OneNotePdfConversionEngine.Convert(notebook, options, cancellationToken);

    /// <summary>Converts the OneNote notebook to a PDF model.</summary>
    public static PdfDocument ToPdf(this OneNoteNotebook notebook, OneNotePdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        notebook.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the OneNote notebook and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this OneNoteNotebook notebook, OneNotePdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = notebook.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, Source, options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the OneNote notebook, returning the model and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this OneNoteNotebook notebook,
        OneNotePdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new OneNotePdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => notebook.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        PdfDocument value = PdfConversionEvidence.CompleteModel(converted, Source, options.FailOnLoss,
            timer, cancellationToken, out OfficeConversionReport report);
        return new OfficeConversionResult<PdfDocument>(value, report);
    }

    /// <summary>Writes the OneNote notebook as a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this OneNoteNotebook notebook, string path,
        OneNotePdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = notebook.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, path, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes the OneNote notebook as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this OneNoteNotebook notebook, Stream stream,
        OneNotePdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = notebook.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, stream, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OneNote notebook as a PDF file and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this OneNoteNotebook notebook, string path,
        OneNotePdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new OneNotePdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => notebook.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OneNote notebook as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this OneNoteNotebook notebook, Stream stream,
        OneNotePdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new OneNotePdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => notebook.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
