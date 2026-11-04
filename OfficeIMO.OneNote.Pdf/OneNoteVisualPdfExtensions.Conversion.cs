using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

public static partial class OneNoteVisualPdfExtensions {
    private const string Source = "OneNote";

    internal static PdfDocumentConversionResult ConvertToVisualPdfWithEvidence(this OneNoteSection section, OneNoteVisualPdfExportOptions? options,
        CancellationToken cancellationToken) =>
        RenderVisualPdf(section, options, cancellationToken);

    /// <summary>Converts the OneNote section as page images to a PDF model.</summary>
    public static PdfDocument ToVisualPdf(this OneNoteSection section, OneNoteVisualPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        section.ToVisualPdf(options, out _, cancellationToken);

    /// <summary>Converts the OneNote section as page images and returns its completed conversion findings.</summary>
    public static PdfDocument ToVisualPdf(this OneNoteSection section, OneNoteVisualPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = section.ConvertToVisualPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, Source, options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the OneNote section as page images, returning the model and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToVisualPdfAsync(this OneNoteSection section,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new OneNoteVisualPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => section.ConvertToVisualPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        PdfDocument value = PdfConversionEvidence.CompleteModel(converted, Source, options.FailOnLoss,
            timer, cancellationToken, out OfficeConversionReport report);
        return new OfficeConversionResult<PdfDocument>(value, report);
    }

    /// <summary>Writes the OneNote section as page images as a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsVisualPdf(this OneNoteSection section, string path,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = section.ConvertToVisualPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, path, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes the OneNote section as page images as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsVisualPdf(this OneNoteSection section, Stream stream,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = section.ConvertToVisualPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, stream, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OneNote section as page images as a PDF file and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsVisualPdfAsync(this OneNoteSection section, string path,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new OneNoteVisualPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => section.ConvertToVisualPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OneNote section as page images as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsVisualPdfAsync(this OneNoteSection section, Stream stream,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new OneNoteVisualPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => section.ConvertToVisualPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    internal static PdfDocumentConversionResult ConvertToVisualPdfWithEvidence(this OneNoteNotebook notebook, OneNoteVisualPdfExportOptions? options,
        CancellationToken cancellationToken) =>
        RenderVisualPdf(notebook, options, cancellationToken);

    /// <summary>Converts the OneNote notebook as page images to a PDF model.</summary>
    public static PdfDocument ToVisualPdf(this OneNoteNotebook notebook, OneNoteVisualPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdf(options, out _, cancellationToken);

    /// <summary>Converts the OneNote notebook as page images and returns its completed conversion findings.</summary>
    public static PdfDocument ToVisualPdf(this OneNoteNotebook notebook, OneNoteVisualPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = notebook.ConvertToVisualPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, Source, options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the OneNote notebook as page images, returning the model and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToVisualPdfAsync(this OneNoteNotebook notebook,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new OneNoteVisualPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => notebook.ConvertToVisualPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        PdfDocument value = PdfConversionEvidence.CompleteModel(converted, Source, options.FailOnLoss,
            timer, cancellationToken, out OfficeConversionReport report);
        return new OfficeConversionResult<PdfDocument>(value, report);
    }

    /// <summary>Writes the OneNote notebook as page images as a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsVisualPdf(this OneNoteNotebook notebook, string path,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = notebook.ConvertToVisualPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, path, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes the OneNote notebook as page images as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsVisualPdf(this OneNoteNotebook notebook, Stream stream,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = notebook.ConvertToVisualPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, stream, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OneNote notebook as page images as a PDF file and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsVisualPdfAsync(this OneNoteNotebook notebook, string path,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new OneNoteVisualPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => notebook.ConvertToVisualPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OneNote notebook as page images as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsVisualPdfAsync(this OneNoteNotebook notebook, Stream stream,
        OneNoteVisualPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new OneNoteVisualPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => notebook.ConvertToVisualPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
