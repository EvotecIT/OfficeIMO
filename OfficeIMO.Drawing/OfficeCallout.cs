namespace OfficeIMO.Drawing;

/// <summary>
/// Positioned callout geometry and plain-text content used by dependency-free renderers.
/// </summary>
public sealed class OfficeCallout {
    /// <summary>
    /// Creates a callout with absolute coordinates in the caller's visual coordinate space.
    /// </summary>
    public OfficeCallout(
        double x,
        double y,
        double width,
        double height,
        double anchorX,
        double anchorY,
        string? title,
        string? text) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        AnchorX = anchorX;
        AnchorY = anchorY;
        Title = title ?? string.Empty;
        Text = text ?? string.Empty;
    }

    /// <summary>X position in CSS pixels before renderer scale is applied.</summary>
    public double X { get; }

    /// <summary>Y position in CSS pixels before renderer scale is applied.</summary>
    public double Y { get; }

    /// <summary>Callout width in CSS pixels before renderer scale is applied.</summary>
    public double Width { get; }

    /// <summary>Callout height in CSS pixels before renderer scale is applied.</summary>
    public double Height { get; }

    /// <summary>X coordinate of the external anchor point before renderer scale is applied.</summary>
    public double AnchorX { get; }

    /// <summary>Y coordinate of the external anchor point before renderer scale is applied.</summary>
    public double AnchorY { get; }

    /// <summary>Header/title text.</summary>
    public string Title { get; }

    /// <summary>Body text.</summary>
    public string Text { get; }
}
