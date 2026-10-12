namespace OfficeIMO.Pdf;

/// <summary>Permission-neutral page geometry used by viewing and print-layout workflows.</summary>
public sealed class PdfPageLayoutInfo {
    internal PdfPageLayoutInfo(
        int pageNumber,
        double width,
        double height,
        double visualWidth,
        double visualHeight,
        int rotationDegrees,
        double userUnit,
        PdfPageGeometry geometry) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        VisualWidth = visualWidth;
        VisualHeight = visualHeight;
        RotationDegrees = rotationDegrees;
        UserUnit = userUnit;
        Geometry = geometry;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Effective CropBox or MediaBox width in default user-space units.</summary>
    public double Width { get; }

    /// <summary>Effective CropBox or MediaBox height in default user-space units.</summary>
    public double Height { get; }

    /// <summary>Displayed page width after rotation and UserUnit are applied.</summary>
    public double VisualWidth { get; }

    /// <summary>Displayed page height after rotation and UserUnit are applied.</summary>
    public double VisualHeight { get; }

    /// <summary>Inherited page rotation normalized to 0, 90, 180, or 270 degrees.</summary>
    public int RotationDegrees { get; }

    /// <summary>Effective positive UserUnit value, defaulting to 1.</summary>
    public double UserUnit { get; }

    /// <summary>Typed page boundary and presentation metadata.</summary>
    public PdfPageGeometry Geometry { get; }

    /// <summary>Maps a top-left visual point to default user space without extracting page content.</summary>
    public PdfPagePoint MapVisualPointToUserSpace(double x, double y) {
        var mapped = PdfVisualCoordinateMapper.TransformVisualPointToUser(GetBoundaryBox(), RotationDegrees, x, y, UserUnit);
        return new PdfPagePoint(mapped.X, mapped.Y);
    }

    /// <summary>Maps a top-left visual rectangle to default user space, accounting for crop, rotation, and user-unit scale.</summary>
    public PdfPageRectangle MapVisualRectangleToUserSpace(double left, double top, double right, double bottom) {
        if (right <= left) throw new ArgumentOutOfRangeException(nameof(right), "Visual rectangle right must be greater than left.");
        if (bottom <= top) throw new ArgumentOutOfRangeException(nameof(bottom), "Visual rectangle bottom must be greater than top.");
        var mapped = PdfVisualCoordinateMapper.TransformVisualBoundsToUser(GetBoundaryBox(), RotationDegrees, left, top, right, bottom, UserUnit);
        return new PdfPageRectangle(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
    }

    /// <summary>Maps a default user-space rectangle to top-left visual coordinates without extracting page content.</summary>
    public PdfSelectionQuad MapUserSpaceRectangleToVisual(double left, double bottom, double right, double top) {
        if (!IsFinite(left)) throw new ArgumentOutOfRangeException(nameof(left));
        if (!IsFinite(bottom)) throw new ArgumentOutOfRangeException(nameof(bottom));
        if (!IsFinite(right) || right <= left) throw new ArgumentOutOfRangeException(nameof(right));
        if (!IsFinite(top) || top <= bottom) throw new ArgumentOutOfRangeException(nameof(top));
        var mapped = PdfVisualCoordinateMapper.TransformBounds(GetBoundaryBox(), RotationDegrees, left, bottom, right, top, UserUnit);
        return new PdfSelectionQuad(new(mapped.Left, mapped.Top), new(mapped.Right, mapped.Top),
            new(mapped.Right, mapped.Bottom), new(mapped.Left, mapped.Bottom));
    }

    private PdfPageBox GetBoundaryBox() {
        if (Geometry.EffectiveBox is { } box) return box;
        if (Geometry.HasEmptyEffectiveBoxIntersection) throw new InvalidOperationException("The page CropBox does not intersect its MediaBox; visual coordinates cannot be mapped safely.");
        return new PdfPageBox("MediaBox", 0, 0, 612, 792);
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
