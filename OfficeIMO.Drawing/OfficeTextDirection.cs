namespace OfficeIMO.Drawing;

/// <summary>Base direction of a Unicode text run.</summary>
public enum OfficeTextDirection {
    /// <summary>No strong directional character was found.</summary>
    Auto = 0,

    /// <summary>The first strong character establishes a left-to-right base direction.</summary>
    LeftToRight = 1,

    /// <summary>The first strong character establishes a right-to-left base direction.</summary>
    RightToLeft = 2
}
