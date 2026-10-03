namespace OfficeIMO.Pdf;

/// <summary>
/// Describes the visible portion of a larger table cell. The renderer lays out
/// text, diagonal borders, data bars and icons in the full box and clips them to
/// the containing cell, preserving wrapping and alignment across page fragments.
/// Images, check boxes and form fields require separate placement.
/// </summary>
public sealed class PdfTableCellViewport {
    /// <summary>
    /// Creates a viewport in one consistent coordinate system. Width and height are
    /// the visible fragment's dimensions; content dimensions describe the unsplit
    /// cell, and offsets locate the fragment from that cell's top-left corner.
    /// The viewport scales with the rendered cell's width and height.
    /// </summary>
    public PdfTableCellViewport(double contentWidth, double contentHeight, double width, double height, double offsetX = 0D, double offsetY = 0D) {
        ValidatePositive(contentWidth, nameof(contentWidth));
        ValidatePositive(contentHeight, nameof(contentHeight));
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        ValidateOffset(offsetX, nameof(offsetX));
        ValidateOffset(offsetY, nameof(offsetY));
        if (width > contentWidth || offsetX > contentWidth - width)
            throw new ArgumentOutOfRangeException(nameof(width), "The horizontal viewport must fit inside the content box.");
        if (height > contentHeight || offsetY > contentHeight - height)
            throw new ArgumentOutOfRangeException(nameof(height), "The vertical viewport must fit inside the content box.");
        if (double.IsInfinity(contentWidth / width) || double.IsInfinity(contentHeight / height))
            throw new ArgumentOutOfRangeException(nameof(width), "Viewport scale ratios must remain finite.");
        ContentWidth = contentWidth;
        ContentHeight = contentHeight;
        Width = width;
        Height = height;
        OffsetX = offsetX;
        OffsetY = offsetY;
    }

    /// <summary>Width of the full content box before clipping.</summary>
    public double ContentWidth { get; }
    /// <summary>Height of the full content box before clipping.</summary>
    public double ContentHeight { get; }
    /// <summary>Width of the visible fragment.</summary>
    public double Width { get; }
    /// <summary>Height of the visible fragment.</summary>
    public double Height { get; }
    /// <summary>Horizontal offset of the fragment within the full content box.</summary>
    public double OffsetX { get; }
    /// <summary>Vertical offset of the fragment within the full content box.</summary>
    public double OffsetY { get; }

    private static void ValidatePositive(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D)
            throw new ArgumentOutOfRangeException(name, "Content and viewport dimensions must be positive finite values.");
    }

    private static void ValidateOffset(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D)
            throw new ArgumentOutOfRangeException(name, "Viewport offsets must be non-negative finite values.");
    }
}
