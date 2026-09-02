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
}
