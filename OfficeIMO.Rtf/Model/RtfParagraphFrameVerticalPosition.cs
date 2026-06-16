namespace OfficeIMO.Rtf;

/// <summary>
/// Vertical placement mode for positioned RTF paragraphs.
/// </summary>
public enum RtfParagraphFrameVerticalPosition {
    /// <summary>Place the frame at an absolute y offset, represented by <c>\posy</c>.</summary>
    Absolute,

    /// <summary>Place the frame at a negative-capable y offset, represented by <c>\posnegy</c>.</summary>
    NegativeAbsolute,

    /// <summary>Place the frame at the top of the vertical reference frame, represented by <c>\posyt</c>.</summary>
    Top,

    /// <summary>Center the frame in the vertical reference frame, represented by <c>\posyc</c>.</summary>
    Center,

    /// <summary>Place the frame at the bottom of the vertical reference frame, represented by <c>\posyb</c>.</summary>
    Bottom,

    /// <summary>Place the frame inline, represented by <c>\posyil</c>.</summary>
    Inline,

    /// <summary>Place the frame inside the vertical reference frame, represented by <c>\posyin</c>.</summary>
    Inside,

    /// <summary>Place the frame outside the vertical reference frame, represented by <c>\posyout</c>.</summary>
    Outside
}
