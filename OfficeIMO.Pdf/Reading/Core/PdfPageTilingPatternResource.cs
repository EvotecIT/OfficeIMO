using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageTilingPatternResource {
    public PdfPageTilingPatternResource(
        OfficeDrawing tile,
        double horizontalStep,
        double verticalStep,
        Matrix2D matrix,
        double boundingBoxX,
        double boundingBoxTop,
        bool uncolored) {
        Tile = tile;
        HorizontalStep = horizontalStep;
        VerticalStep = verticalStep;
        Matrix = matrix;
        BoundingBoxX = boundingBoxX;
        BoundingBoxTop = boundingBoxTop;
        Uncolored = uncolored;
    }

    public OfficeDrawing Tile { get; }
    public double HorizontalStep { get; }
    public double VerticalStep { get; }
    public Matrix2D Matrix { get; }
    public double BoundingBoxX { get; }
    public double BoundingBoxTop { get; }
    public bool Uncolored { get; }
}

internal sealed class PdfPageTilingPatternPaint {
    public PdfPageTilingPatternPaint(PdfPageTilingPatternResource resource, OfficeTransform transform, OfficeColor? tint, double opacity) {
        Resource = resource;
        Transform = transform;
        Tint = tint;
        Opacity = opacity;
    }

    public PdfPageTilingPatternResource Resource { get; }
    public OfficeTransform Transform { get; }
    public OfficeColor? Tint { get; }
    public double Opacity { get; }
}
