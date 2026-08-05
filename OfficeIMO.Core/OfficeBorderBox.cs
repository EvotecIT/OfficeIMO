namespace OfficeIMO.Drawing;

/// <summary>
/// Describes the four independent edges of a rectangular border box.
/// </summary>
public readonly struct OfficeBorderBox {
    /// <summary>
    /// Creates a border box with independent edges.
    /// </summary>
    public OfficeBorderBox(
        OfficeBorderSide? left = null,
        OfficeBorderSide? top = null,
        OfficeBorderSide? right = null,
        OfficeBorderSide? bottom = null,
        OfficeBorderSide? diagonalDown = null,
        OfficeBorderSide? diagonalUp = null) {
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
        DiagonalDown = diagonalDown;
        DiagonalUp = diagonalUp;
    }

    /// <summary>Left edge, or null when absent.</summary>
    public OfficeBorderSide? Left { get; }

    /// <summary>Top edge, or null when absent.</summary>
    public OfficeBorderSide? Top { get; }

    /// <summary>Right edge, or null when absent.</summary>
    public OfficeBorderSide? Right { get; }

    /// <summary>Bottom edge, or null when absent.</summary>
    public OfficeBorderSide? Bottom { get; }

    /// <summary>Top-left to bottom-right diagonal edge, or null when absent.</summary>
    public OfficeBorderSide? DiagonalDown { get; }

    /// <summary>Bottom-left to top-right diagonal edge, or null when absent.</summary>
    public OfficeBorderSide? DiagonalUp { get; }

    /// <summary>Whether at least one edge is visible.</summary>
    public bool HasVisibleSide =>
        (Left?.IsVisible ?? false) ||
        (Top?.IsVisible ?? false) ||
        (Right?.IsVisible ?? false) ||
        (Bottom?.IsVisible ?? false) ||
        (DiagonalDown?.IsVisible ?? false) ||
        (DiagonalUp?.IsVisible ?? false);

    /// <summary>
    /// Creates a border box where all edges use the same side definition.
    /// </summary>
    public static OfficeBorderBox Uniform(OfficeBorderSide side) => new OfficeBorderBox(side, side, side, side);
}
