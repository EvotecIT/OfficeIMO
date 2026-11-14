using System.Diagnostics;
using OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Converts the source document to a PDF model.</summary>
    public static PdfDocument ToPdf(this WordDocument document, WordPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        document.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the source document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this WordDocument document, WordPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using IDisposable layoutScope = UseDeferredLayoutScope(document, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, "Word", options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the source document, returning the model and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this WordDocument document,
        WordPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordPdfExportOptions()).Clone();
        return await Task.Run(() => {
            PdfDocument value = document.ToPdf(options, out OfficeConversionReport report, cancellationToken);
            try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this WordDocument document, string path,
        WordPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new WordPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        using IDisposable layoutScope = UseDeferredLayoutScope(document, cancellationToken);
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "Word", path, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "Word", () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport("Word", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this WordDocument document, Stream stream,
        WordPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        using IDisposable layoutScope = UseDeferredLayoutScope(document, cancellationToken);
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, "Word", stream, timer, cancellationToken) : PdfConversionEvidence.RunOutput(converted, "Word", () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport("Word", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes a PDF file and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this WordDocument document, string path,
        WordPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new WordPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => document.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        using IDisposable layoutScope = UseDeferredLayoutScope(document, cancellationToken);
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Word", path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "Word", () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Word", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes PDF to a caller-owned stream and returns its conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this WordDocument document, Stream stream,
        WordPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new WordPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => document.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        using IDisposable layoutScope = UseDeferredLayoutScope(document, cancellationToken);
        PdfSaveResult saved = options?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Word", stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, "Word", () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Word", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>
    /// Header, footer and paragraph callbacks run while the PDF model lays out pages, after
    /// the conversion's own scopes have ended. Without them every list lookup in those
    /// callbacks rebuilds the styles and numbering catalog, once per page and paragraph, and
    /// run paint resolution walks table and style chains again for each run. The document is
    /// not edited during layout, so one catalog and one run paint cache serve the whole pass.
    /// </summary>
    private static IDisposable UseDeferredLayoutScope(WordDocument document, CancellationToken cancellationToken) =>
        new DeferredLayoutScope(
            WordRunFormatting.BeginConversionScope(cancellationToken),
            WordDocumentTraversal.UseListNumberingCatalog(document));

    private sealed class DeferredLayoutScope : IDisposable {
        private readonly IDisposable _runPaint;
        private readonly IDisposable _listCatalog;

        internal DeferredLayoutScope(IDisposable runPaint, IDisposable listCatalog) {
            _runPaint = runPaint;
            _listCatalog = listCatalog;
        }

        public void Dispose() {
            _listCatalog.Dispose();
            _runPaint.Dispose();
        }
    }
}
