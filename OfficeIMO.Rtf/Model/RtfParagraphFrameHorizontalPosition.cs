namespace OfficeIMO.Rtf;

/// <summary>
/// Horizontal placement mode for positioned RTF paragraphs.
/// </summary>
public enum RtfParagraphFrameHorizontalPosition {
    /// <summary>Place the frame at an absolute x offset, represented by <c>\posx</c>.</summary>
    Absolute,

    /// <summary>Place the frame at a negative-capable x offset, represented by <c>\posnegx</c>.</summary>
    NegativeAbsolute,

    /// <summary>Place the frame at the left of the horizontal reference frame, represented by <c>\posxl</c>.</summary>
    Left,

    /// <summary>Center the frame in the horizontal reference frame, represented by <c>\posxc</c>.</summary>
    Center,

    /// <summary>Place the frame at the right of the horizontal reference frame, represented by <c>\posxr</c>.</summary>
    Right,

    /// <summary>Place the frame inside the horizontal reference frame, represented by <c>\posxi</c>.</summary>
    Inside,

    /// <summary>Place the frame outside the horizontal reference frame, represented by <c>\posxo</c>.</summary>
    Outside
}
