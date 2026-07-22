using System.Diagnostics;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageImageRenderer {
    /// <summary>Renders all pages or a caller-ordered page selection with bounded per-page reports.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        PdfPageSelection? selection = null,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfPageRenderOptions effectiveOptions = options ?? new PdfPageRenderOptions();
        effectiveOptions.Validate();
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        int[] pages = selection?.ToPageNumbers(document.Pages.Count, nameof(selection)) ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        if (pages.Length > effectiveOptions.MaxPages) {
            throw new PdfReadLimitException(PdfReadLimitKind.RenderPages, effectiveOptions.MaxPages, pages.Length, "PDF render page count exceeded the configured limit.");
        }

        var results = new List<PdfPageRenderResult>(pages.Length);
        long totalOutputBytes = 0;
        for (int i = 0; i < pages.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfPageRenderResult result = RenderPage(document, pages[i], effectiveOptions, cancellationToken);
            totalOutputBytes = checked(totalOutputBytes + result.OutputByteLength);
            if (totalOutputBytes > effectiveOptions.MaxTotalOutputBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effectiveOptions.MaxTotalOutputBytes, totalOutputBytes);
            }
            results.Add(result);
        }

        return results.AsReadOnly();
    }

    /// <summary>Renders parsed page ranges such as <c>1-3,5</c>.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        string pageRanges,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        return RenderPages(pdf, PdfPageSelector.Parse(pageRanges).ResolveSelection(document.Pages.Count), options, readOptions, cancellationToken);
    }

    /// <summary>Renders pages resolved by a document-relative selector.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        PdfPageSelector selector,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(selector, nameof(selector));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        return RenderPages(pdf, selector.ResolveSelection(document.Pages.Count), options, readOptions, cancellationToken);
    }

    private static PdfPageRenderResult RenderPage(PdfReadDocument document, int pageNumber, PdfPageRenderOptions options, CancellationToken cancellationToken) {
        var timer = Stopwatch.StartNew();
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics = Array.Empty<PdfRenderCapabilityDiagnostic>();
        try {
            cancellationToken.ThrowIfCancellationRequested();
            capabilityDiagnostics = document.Pages[pageNumber - 1].GetRenderCapabilityDiagnostics();
            OfficeDrawing drawing = RenderPage(document, pageNumber);
            double scale = options.GetScale(drawing);
            int width = checked((int)Math.Ceiling(drawing.Width * scale));
            int height = checked((int)Math.Ceiling(drawing.Height * scale));
            long pixels = checked((long)width * height);
            if (pixels > options.MaxPixelsPerPage) {
                throw new PdfReadLimitException(PdfReadLimitKind.RenderPixels, options.MaxPixelsPerPage, pixels, "PDF render pixel count exceeded the configured per-page limit.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = options.Format == PdfPageRenderFormat.Png
                ? RenderDrawingAsPng(drawing, scale, options.Background, options.ImageCodec)
                : OfficeDrawingSvgExporter.ToSvgBytes(drawing, scale);
            if (bytes.LongLength > options.MaxOutputBytesPerPage) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, options.MaxOutputBytesPerPage, bytes.LongLength);
            }
            timer.Stop();
            return new PdfPageRenderResult(pageNumber, options.Format, bytes, width, height, timer.Elapsed, capabilityDiagnostics);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) when (options.ContinueOnError && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            timer.Stop();
            return new PdfPageRenderResult(pageNumber, options.Format, null, 0, 0, timer.Elapsed, capabilityDiagnostics, new[] { ex.GetType().Name + ": " + ex.Message });
        }
    }
}
