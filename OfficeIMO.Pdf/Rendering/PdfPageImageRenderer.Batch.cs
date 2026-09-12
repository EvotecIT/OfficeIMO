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
        return RenderPages(token => PdfReadDocument.Open(pdf, readOptions, token), selection, options, cancellationToken);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfPageRenderResult> RenderPagesCore(
        Func<CancellationToken, PdfReadDocument> getDocument,
        Func<int, int[]> resolvePages,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        PdfPageRenderOptions effectiveOptions = options ?? new PdfPageRenderOptions();
        effectiveOptions.Validate();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effectiveOptions.RenderTimeout,
            cancellationToken);
        try {
            execution.Token.ThrowIfCancellationRequested();
            PdfReadDocument document = getDocument(execution.Token);
            Guard.NotNull(document, nameof(document));
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
        return RenderPages(token => PdfReadDocument.Open(pdf, readOptions, token), pageRanges, options, cancellationToken);
    }

    /// <summary>Renders pages resolved by a document-relative selector.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        PdfPageSelector selector,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(token => PdfReadDocument.Open(pdf, readOptions, token), selector, options, cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<CancellationToken, PdfReadDocument> getDocument,
        PdfPageSelection? selection,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(getDocument, nameof(getDocument));
        return RenderPagesCore(
            getDocument,
            pageCount => selection?.ToPageNumbers(pageCount, nameof(selection)) ?? Enumerable.Range(1, pageCount).ToArray(),
            options,
            cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<CancellationToken, PdfReadDocument> getDocument,
        string pageRanges,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(getDocument, nameof(getDocument));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return RenderPagesCore(
            getDocument,
            pageCount => PdfPageSelector.Parse(pageRanges).ResolveSelection(pageCount).ToPageNumbers(pageCount, nameof(pageRanges)),
            options,
            cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<CancellationToken, PdfReadDocument> getDocument,
        PdfPageSelector selector,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(getDocument, nameof(getDocument));
        Guard.NotNull(selector, nameof(selector));
        return RenderPagesCore(
            getDocument,
            pageCount => selector.ResolveSelection(pageCount).ToPageNumbers(pageCount, nameof(selector)),
            options,
            cancellationToken);
    }

    internal static PdfPageRenderResult RenderPage(PdfReadDocument document, int pageNumber, PdfPageRenderOptions options, CancellationToken cancellationToken, bool forDisplay = false) {
        var timer = Stopwatch.StartNew();
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics = Array.Empty<PdfRenderCapabilityDiagnostic>();
        try {
            cancellationToken.ThrowIfCancellationRequested();
            capabilityDiagnostics = document.Pages[pageNumber - 1].GetRenderCapabilityDiagnostics(cancellationToken);
            void ConfigureDrawing(OfficeDrawing scene) {
                scene.Fonts.AddRangePreservingExisting(options.Fonts);
                scene.TextShapingProvider = options.TextShapingProvider;
                scene.TextShapingLanguage = options.TextShapingLanguage;
            }
            OfficeDrawing drawing = forDisplay ? document.Pages[pageNumber - 1].ToDisplayDrawing(cancellationToken, ConfigureDrawing)
                : document.Pages[pageNumber - 1].ToDrawing(cancellationToken, ConfigureDrawing);
            double scale = options.GetScale(drawing);
            int width = checked((int)Math.Ceiling(drawing.Width * scale));
            int height = checked((int)Math.Ceiling(drawing.Height * scale));
            long pixels = checked((long)width * height);
            if (pixels > options.MaxPixelsPerPage) {
                throw new PdfReadLimitException(PdfReadLimitKind.RenderPixels, options.MaxPixelsPerPage, pixels, "PDF render pixel count exceeded the configured per-page limit.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes;
            if (options.Format == PdfPageRenderFormat.Png) {
                bytes = RenderDrawingAsPng(
                    drawing,
                    scale,
                    options.Background,
                    options.ImageCodec,
                    options.MaxPixelsPerPage,
                    options.TextShapingProvider,
                    options.TextShapingLanguage,
                    cancellationToken);
            } else {
                try {
                    bytes = OfficeDrawingSvgExporter.ToSvgBytes(
                        drawing,
                        scale,
                        OfficeSvgSizeUnit.Point,
                        imageCodec: null,
                        resourceIdPrefix: null,
                        maximumUtf8Bytes: options.MaxOutputBytesPerPage,
                        cancellationToken: cancellationToken);
                } catch (OfficeImageExportBatchLimitException exception) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.RenderBytes,
                        options.MaxOutputBytesPerPage,
                        exception.Actual);
                }
            }
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
