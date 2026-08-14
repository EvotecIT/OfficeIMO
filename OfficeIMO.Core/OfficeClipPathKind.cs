namespace OfficeIMO.Drawing;

/// <summary>Supported reusable clipping path descriptors.</summary>
public enum OfficeClipPathKind {
    /// <summary>Rectangular clipping path.</summary>
    Rectangle = 0,

    /// <summary>Rounded rectangle clipping path.</summary>
    RoundedRectangle = 1,

    /// <summary>Freeform path clipping descriptor.</summary>
    Path = 2,

    /// <summary>Empty clipping region that suppresses paint while retaining nested content semantics.</summary>
    Empty = 3
}
