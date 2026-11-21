namespace OfficeIMO.Pdf;

/// <summary>
/// Lays out ordinary flow content inside a fixed rectangle, as a positioned text frame does.
/// The frame paints at its rectangle and takes no room in the surrounding flow; content outside
/// the clip rectangle is not painted.
/// </summary>
internal sealed class PdfCanvasFlowItem : PdfCanvasItem {
    internal PdfCanvasFlowItem(IReadOnlyList<IPdfBlock> blocks, double x, double y, double width, double height,
        double clipX, double clipY, double clipWidth, double clipHeight) : base(x, y) {
        Blocks = blocks;
        Width = width;
        Height = height;
        ClipX = clipX;
        ClipY = clipY;
        ClipWidth = clipWidth;
        ClipHeight = clipHeight;
    }

    internal IReadOnlyList<IPdfBlock> Blocks { get; }
    internal double Width { get; }
    internal double Height { get; }
    internal double ClipX { get; }
    internal double ClipY { get; }
    internal double ClipWidth { get; }
    internal double ClipHeight { get; }
}

public sealed partial class PdfPageCanvas {
    /// <summary>
    /// Adds flow blocks laid out in the content rectangle (measured from the page top) and clipped to
    /// the outer rectangle. The surrounding flow keeps its position.
    /// </summary>
    internal PdfPageCanvas PlacedFlow(IReadOnlyList<IPdfBlock> blocks, double x, double y, double width, double height,
        double clipX, double clipY, double clipWidth, double clipHeight) {
        Guard.NotNull(blocks, nameof(blocks));
        if (!IsFinite(x) || !IsFinite(y) || !IsFinite(clipX) || !IsFinite(clipY) ||
            !(width > 0D) || !(height > 0D) || !(clipWidth > 0D) || !(clipHeight > 0D) ||
            !IsFinite(width) || !IsFinite(height) || !IsFinite(clipWidth) || !IsFinite(clipHeight))
            throw new ArgumentOutOfRangeException(nameof(width), "A placed flow needs finite, positive rectangles.");
        _items.Add(new PdfCanvasFlowItem(blocks, x, y, width, height, clipX, clipY, clipWidth, clipHeight));
        return this;

        static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
