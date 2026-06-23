namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free hatch pattern primitives used by OfficeIMO image renderers.
/// </summary>
public enum OfficeHatchPatternKind {
    /// <summary>Horizontal hatch lines.</summary>
    Horizontal,

    /// <summary>Vertical hatch lines.</summary>
    Vertical,

    /// <summary>Diagonal lines descending from left to right.</summary>
    DiagonalDown,

    /// <summary>Diagonal lines ascending from left to right.</summary>
    DiagonalUp,

    /// <summary>Horizontal and vertical hatch lines.</summary>
    Grid,

    /// <summary>Crossed diagonal hatch lines.</summary>
    Trellis,

    /// <summary>Small repeated square dots.</summary>
    Dotted
}
