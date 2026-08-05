namespace OfficeIMO.Drawing;

/// <summary>
/// Defines the unit written on the root SVG width and height attributes.
/// The SVG view box always remains in drawing-local coordinates.
/// </summary>
public enum OfficeSvgSizeUnit {
    /// <summary>Write physical point dimensions for document-oriented drawing consumers.</summary>
    Point,

    /// <summary>Write CSS pixel dimensions for screen and image-oriented consumers.</summary>
    Pixel
}
