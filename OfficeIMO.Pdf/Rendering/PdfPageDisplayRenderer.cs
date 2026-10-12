using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Renders rectangles of pages to pixels for an interactive viewer. Recently used page drawings and their decoded and
/// prefiltered images are kept between renders, so the tiles, zoom levels and repeated views of a page read and decode
/// it once. As with <see cref="PdfDocumentRenderer.DisplayPage"/>, only pixels leave it, so documents that restrict
/// content copying render the same way. One instance serves one unchanged document; callers serialize its renders.
/// </summary>
internal sealed class PdfPageDisplayRenderer {
    private readonly PdfDocument _document;
    private readonly int _pageCapacity;
    private readonly LinkedList<CachedDrawing> _drawings = new();
    private readonly OfficeRasterImageCache _images;

    internal PdfPageDisplayRenderer(PdfDocument document, int pageCapacity, long imageCapacityBytes) {
        Guard.NotNull(document, nameof(document));
        _document = document;
        if (pageCapacity < 1) throw new ArgumentOutOfRangeException(nameof(pageCapacity), "Keep at least one page.");
        _pageCapacity = pageCapacity;
        _images = new OfficeRasterImageCache(imageCapacityBytes);
    }

    /// <summary>Decoded image bytes currently kept.</summary>
    internal long CachedImageBytes => _images.UsedBytes;

    /// <summary>
    /// Renders the pixel rectangle <paramref name="x"/>, <paramref name="y"/>, <paramref name="width"/> by
    /// <paramref name="height"/> of the page at <paramref name="scale"/> pixels per point, on white. The page image at
    /// that scale is its visual size times the scale, each side rounded up; tiles of it join without seams.
    /// </summary>
    internal OfficeRasterImage RenderRegion(int pageNumber, double scale, int x, int y, int width, int height,
        PdfPageRegionLimits limits, CancellationToken cancellationToken) {
        Guard.NotNull(limits, nameof(limits));
        if (!(scale > 0D) || double.IsInfinity(scale))
            throw new ArgumentOutOfRangeException(nameof(scale), "The scale must be a finite positive number.");
        long pixels = (long)Math.Max(0, width) * Math.Max(0, height);
        if (pixels > limits.MaximumRegionPixels) {
            throw new PdfReadLimitException(PdfReadLimitKind.RenderPixels, limits.MaximumRegionPixels, pixels,
                "The requested region exceeds the per-render pixel limit.");
        }
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(limits.Timeout, cancellationToken);
        try {
            execution.Token.ThrowIfCancellationRequested();
            PdfReadDocument document = _document.GetReadDocument(_document.ReadOptions, execution.Token);
            if (pageNumber <= 0 || pageNumber > document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must refer to an existing one-based PDF page.");
            }
            OfficeDrawing drawing = GetDrawing(document, pageNumber, ShadingScale(scale), limits.MaximumWorkingPixels, execution.Token);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.RenderRegion(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = scale,
                Background = OfficeColor.White,
                ThrowOnImageDecodeFailure = true,
                MaximumRasterPixels = limits.MaximumWorkingPixels,
                CancellationToken = execution.Token,
                ImageCache = _images
            }, x, y, width, height);
            execution.ThrowIfCancellationRequested();
            return image;
        } catch (OperationCanceledException error) when (execution.IsTimeoutCancellation(error)) {
            throw execution.CreateTimeoutException(error);
        }
    }

    // Function shadings are sampled for the scale they are drawn at. Drawings are kept per power of two of the scale,
    // and a drawing sampled for a larger scale also serves smaller ones.
    private static double ShadingScale(double scale) {
        double shading = 1D;
        while (shading < scale && shading < 16D) shading *= 2D;
        return shading;
    }

    private OfficeDrawing GetDrawing(PdfReadDocument document, int pageNumber, double shadingScale, long maximumPixels,
        CancellationToken cancellationToken) {
        for (LinkedListNode<CachedDrawing>? node = _drawings.First; node != null; node = node.Next) {
            if (node.Value.PageNumber == pageNumber && node.Value.ShadingScale >= shadingScale) {
                _drawings.Remove(node);
                _drawings.AddFirst(node);
                return node.Value.Drawing;
            }
        }
        OfficeDrawing drawing = document.Pages[pageNumber - 1].ToDisplayDrawing(cancellationToken,
            functionShadingScale: shadingScale, maximumFunctionPixels: maximumPixels);
        for (LinkedListNode<CachedDrawing>? node = _drawings.First; node != null;) {
            LinkedListNode<CachedDrawing>? next = node.Next;
            if (node.Value.PageNumber == pageNumber) _drawings.Remove(node);
            node = next;
        }
        _drawings.AddFirst(new CachedDrawing(pageNumber, shadingScale, drawing));
        while (_drawings.Count > _pageCapacity) _drawings.RemoveLast();
        return drawing;
    }

    private sealed class CachedDrawing {
        internal CachedDrawing(int pageNumber, double shadingScale, OfficeDrawing drawing) {
            PageNumber = pageNumber;
            ShadingScale = shadingScale;
            Drawing = drawing;
        }

        internal int PageNumber { get; }
        internal double ShadingScale { get; }
        internal OfficeDrawing Drawing { get; }
    }
}

/// <summary>Limits applied to one <see cref="PdfPageDisplayRenderer.RenderRegion"/> call.</summary>
internal sealed class PdfPageRegionLimits {
    /// <summary>Largest region, in pixels.</summary>
    internal long MaximumRegionPixels { get; init; } = 16_000_000L;
    /// <summary>The renderer's working budget: the region, intermediate surfaces and decoded source images.</summary>
    internal long MaximumWorkingPixels { get; init; } = OfficeImageExportOptions.DefaultMaximumRasterPixels;
    /// <summary>Deadline for reading the page and rasterizing the region; infinite by default.</summary>
    internal TimeSpan Timeout { get; init; } = System.Threading.Timeout.InfiniteTimeSpan;
}
