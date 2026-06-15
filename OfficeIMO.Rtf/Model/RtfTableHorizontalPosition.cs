namespace OfficeIMO.Rtf;

/// <summary>
/// Horizontal placement mode for a positioned RTF table row.
/// </summary>
public enum RtfTableHorizontalPosition {
    /// <summary>Place the table at an absolute x offset, represented by <c>\tposx</c>.</summary>
    Absolute,

    /// <summary>Place the table at a negative-capable x offset, represented by <c>\tposnegx</c>.</summary>
    NegativeAbsolute,

    /// <summary>Place the table at the left of the horizontal frame, represented by <c>\tposxl</c>.</summary>
    Left,

    /// <summary>Center the table in the horizontal frame, represented by <c>\tposxc</c>.</summary>
    Center,

    /// <summary>Place the table at the right of the horizontal frame, represented by <c>\tposxr</c>.</summary>
    Right,

    /// <summary>Place the table inside the horizontal frame, represented by <c>\tposxi</c>.</summary>
    Inside,

    /// <summary>Place the table outside the horizontal frame, represented by <c>\tposxo</c>.</summary>
    Outside
}
