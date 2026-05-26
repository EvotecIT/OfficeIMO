namespace OfficeIMO.Drawing;

/// <summary>
/// Shared shape kinds that OfficeIMO packages can map into their own document formats.
/// </summary>
public enum OfficeShapeKind {
    /// <summary>Axis-aligned rectangle.</summary>
    Rectangle,

    /// <summary>Axis-aligned ellipse bounded by width and height.</summary>
    Ellipse,

    /// <summary>Closed polygon described by local points.</summary>
    Polygon,

    /// <summary>Freeform path described by path commands.</summary>
    Path,

    /// <summary>Open straight line described by two local points.</summary>
    Line,

    /// <summary>Axis-aligned rectangle with rounded corners.</summary>
    RoundedRectangle
}
