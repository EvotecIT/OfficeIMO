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
    Dotted,

    /// <summary>Stipple fill with roughly 6.25 percent foreground coverage.</summary>
    Percent6_25,

    /// <summary>Stipple fill with roughly 12.5 percent foreground coverage.</summary>
    Percent12_5,

    /// <summary>Stipple fill with roughly 25 percent foreground coverage.</summary>
    Percent25,

    /// <summary>Stipple fill with roughly 50 percent foreground coverage.</summary>
    Percent50,

    /// <summary>Stipple fill with roughly 75 percent foreground coverage.</summary>
    Percent75
}
