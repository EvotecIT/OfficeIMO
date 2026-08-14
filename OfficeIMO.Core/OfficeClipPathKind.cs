namespace OfficeIMO.Drawing;

/// <summary>Supported reusable clipping path descriptors.</summary>
public enum OfficeClipPathKind {
    /// <summary>Empty clipping region that suppresses paint while retaining nested content semantics.</summary>
    Empty,

    /// <summary>Rectangular clipping path.</summary>
    Rectangle,

    /// <summary>Rounded rectangle clipping path.</summary>
    RoundedRectangle,

    /// <summary>Freeform path clipping descriptor.</summary>
    Path
}
