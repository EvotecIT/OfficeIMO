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
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(() => pdf, selection, options, readOptions, cancellationToken);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfPageRenderResult> RenderPagesCore(
        Func<byte[]> getPdf,
        Func<int, int[]> resolvePages,
        PdfPageRenderOptions? options,
        PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        PdfPageRenderOptions effectiveOptions = options ?? new PdfPageRenderOptions();
        effectiveOptions.Validate();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effectiveOptions.RenderTimeout,
            cancellationToken);
        try {
            execution.Token.ThrowIfCancellationRequested();
            byte[] pdf = getPdf();
            Guard.NotNull(pdf, nameof(pdf));
            execution.Token.ThrowIfCancellationRequested();
            PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
            execution.Token.ThrowIfCancellationRequested();
            int[] pages = resolvePages(document.Pages.Count);
            execution.Token.ThrowIfCancellationRequested();
            if (pages.Length > effectiveOptions.MaxPages) {
                throw new PdfReadLimitException(PdfReadLimitKind.RenderPages, effectiveOptions.MaxPages, pages.Length, "PDF render page count exceeded the configured limit.");
            }

            var results = new List<PdfPageRenderResult>(pages.Length);
            long totalOutputBytes = 0;
            for (int i = 0; i < pages.Length; i++) {
                execution.Token.ThrowIfCancellationRequested();
                PdfPageRenderResult result = RenderPage(document, pages[i], effectiveOptions, execution.Token);
                totalOutputBytes = checked(totalOutputBytes + result.OutputByteLength);
                if (totalOutputBytes > effectiveOptions.MaxTotalOutputBytes) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effectiveOptions.MaxTotalOutputBytes, totalOutputBytes);
                }
                results.Add(result);
            }

            execution.ThrowIfCancellationRequested();
            return results.AsReadOnly();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>Renders parsed page ranges such as <c>1-3,5</c>.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        string pageRanges,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(() => pdf, pageRanges, options, readOptions, cancellationToken);
    }

    /// <summary>Renders pages resolved by a document-relative selector.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        PdfPageSelector selector,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(() => pdf, selector, options, readOptions, cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<byte[]> getPdf,
        PdfPageSelection? selection,
        PdfPageRenderOptions? options,
        PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(getPdf, nameof(getPdf));
        return RenderPagesCore(
            getPdf,
            pageCount => selection?.ToPageNumbers(pageCount, nameof(selection)) ?? Enumerable.Range(1, pageCount).ToArray(),
            options,
            readOptions,
            cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<byte[]> getPdf,
        string pageRanges,
        PdfPageRenderOptions? options,
        PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(getPdf, nameof(getPdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return RenderPagesCore(
            getPdf,
            pageCount => PdfPageSelector.Parse(pageRanges).ResolveSelection(pageCount).ToPageNumbers(pageCount, nameof(pageRanges)),
            options,
            readOptions,
            cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<byte[]> getPdf,
        PdfPageSelector selector,
        PdfPageRenderOptions? options,
        PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(getPdf, nameof(getPdf));
        Guard.NotNull(selector, nameof(selector));
        return RenderPagesCore(
            getPdf,
            pageCount => selector.ResolveSelection(pageCount).ToPageNumbers(pageCount, nameof(selector)),
            options,
            readOptions,
            cancellationToken);
    }

    private static PdfPageRenderResult RenderPage(PdfReadDocument document, int pageNumber, PdfPageRenderOptions options, CancellationToken cancellationToken) {
        var timer = Stopwatch.StartNew();
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics = Array.Empty<PdfRenderCapabilityDiagnostic>();
        try {
            cancellationToken.ThrowIfCancellationRequested();
            capabilityDiagnostics = document.Pages[pageNumber - 1].GetRenderCapabilityDiagnostics();
            OfficeDrawing drawing = RenderPage(document, pageNumber);
            drawing.Fonts.AddRangePreservingExisting(options.Fonts);
            double scale = options.GetScale(drawing);
            int width = checked((int)Math.Ceiling(drawing.Width * scale));
            int height = checked((int)Math.Ceiling(drawing.Height * scale));
            long pixels = checked((long)width * height);
            if (pixels > options.MaxPixelsPerPage) {
                throw new PdfReadLimitException(PdfReadLimitKind.RenderPixels, options.MaxPixelsPerPage, pixels, "PDF render pixel count exceeded the configured per-page limit.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = options.Format == PdfPageRenderFormat.Png
                ? RenderDrawingAsPng(
                    drawing,
                    scale,
                    options.Background,
                    options.ImageCodec,
                    options.MaxPixelsPerPage,
                    options.TextShapingProvider,
                    options.TextShapingLanguage,
                    cancellationToken)
                : OfficeDrawingSvgExporter.ToSvgBytes(
                    drawing,
                    scale,
                    OfficeSvgSizeUnit.Point,
                    imageCodec: null,
                    resourceIdPrefix: null,
                    cancellationToken: cancellationToken);
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
