namespace OfficeIMO.Rtf;

/// <summary>
/// Vertical placement mode for a positioned RTF table row.
/// </summary>
public enum RtfTableVerticalPosition {
    /// <summary>Place the table at an absolute y offset, represented by <c>\tposy</c>.</summary>
    Absolute,

    /// <summary>Place the table at a negative-capable y offset, represented by <c>\tposnegy</c>.</summary>
    NegativeAbsolute,

    /// <summary>Place the table at the top of the vertical frame, represented by <c>\tposyt</c>.</summary>
    Top,

    /// <summary>Center the table in the vertical frame, represented by <c>\tposyc</c>.</summary>
    Center,

    /// <summary>Place the table at the bottom of the vertical frame, represented by <c>\tposyb</c>.</summary>
    Bottom,

    /// <summary>Place the table inline, represented by <c>\tposyil</c>.</summary>
    Inline,

    /// <summary>Place the table inside the vertical frame, represented by <c>\tposyin</c>.</summary>
    Inside,

    /// <summary>Place the table outside the vertical frame, represented by <c>\tposyoutv</c>.</summary>
    Outside
}
