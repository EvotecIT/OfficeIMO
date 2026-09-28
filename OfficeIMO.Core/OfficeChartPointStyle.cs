using System;

namespace OfficeIMO.Drawing;

/// <summary>Bounded hatch patterns shared by native chart writers and static renderers.</summary>
public enum OfficeChartHatchPattern {
    /// <summary>Horizontal strokes.</summary>
    Horizontal,
    /// <summary>Vertical strokes.</summary>
    Vertical,
    /// <summary>Strokes rising from left to right.</summary>
    ForwardDiagonal,
    /// <summary>Strokes falling from left to right.</summary>
    BackwardDiagonal,
    /// <summary>Horizontal and vertical strokes.</summary>
    Cross,
    /// <summary>Crossing diagonal strokes.</summary>
    DiagonalCross,
    /// <summary>Widely spaced strokes rising from left to right.</summary>
    WideForwardDiagonal
}

/// <summary>
/// Immutable point appearance. Unspecified properties inherit the chart appearance;
/// explicit no-fill and hidden outlines are distinct from inheritance.
/// </summary>
public sealed class OfficeChartPointStyle {
    /// <summary>Creates a point style with the original fill, hatch, and outline options.</summary>
    public OfficeChartPointStyle(OfficeColor? fillColor, bool noFill,
        OfficeChartHatchPattern? hatch, OfficeColor? hatchColor,
        OfficeColor? outlineColor, double? outlineWidth, bool? showOutline)
        : this(fillColor, noFill, hatch, hatchColor, outlineColor, outlineWidth, showOutline, null) {
    }

    /// <summary>Creates a point style with optional solid fill, hatch, or outline overrides.</summary>
    /// <param name="fillColor">Solid fill, or the hatch background. Null inherits the point colour for solid fills and uses white for hatches.</param>
    /// <param name="noFill">Suppress the fill instead of inheriting a colour.</param>
    /// <param name="hatch">Optional hatch; its strokes are clipped to the point geometry.</param>
    /// <param name="hatchColor">Hatch stroke colour; required when a hatch is specified.</param>
    /// <param name="outlineColor">Optional outline colour.</param>
    /// <param name="outlineWidth">Optional finite positive outline width, at most 1584 points (the DrawingML limit).</param>
    /// <param name="showOutline">Null inherits, false removes the outline, true enables it.</param>
    /// <param name="outlineJoin">Optional outline corner join.</param>
    public OfficeChartPointStyle(OfficeColor? fillColor = null, bool noFill = false,
        OfficeChartHatchPattern? hatch = null, OfficeColor? hatchColor = null,
        OfficeColor? outlineColor = null, double? outlineWidth = null, bool? showOutline = null,
        OfficeStrokeLineJoin? outlineJoin = null) {
        if (noFill && (fillColor.HasValue || hatch.HasValue))
            throw new ArgumentException("No-fill cannot be combined with a solid or hatch fill.", nameof(noFill));
        if (hatch.HasValue != hatchColor.HasValue)
            throw new ArgumentException("A hatch and its stroke colour must be specified together.", nameof(hatchColor));
        if (hatch.HasValue && !Enum.IsDefined(typeof(OfficeChartHatchPattern), hatch.Value))
            throw new ArgumentOutOfRangeException(nameof(hatch));
        if (outlineJoin.HasValue && !Enum.IsDefined(typeof(OfficeStrokeLineJoin), outlineJoin.Value))
            throw new ArgumentOutOfRangeException(nameof(outlineJoin));
        OfficeChartStyleBounds.ValidateLineWidth(outlineWidth, nameof(outlineWidth));
        FillColor = fillColor ?? (hatch.HasValue ? OfficeColor.White : (OfficeColor?)null);
        NoFill = noFill;
        Hatch = hatch;
        HatchColor = hatchColor;
        OutlineColor = outlineColor;
        OutlineWidth = outlineWidth;
        ShowOutline = showOutline;
        OutlineJoin = outlineJoin;
    }

    /// <summary>Solid fill or hatch background override.</summary>
    public OfficeColor? FillColor { get; }
    /// <summary>Whether the point has explicitly no fill.</summary>
    public bool NoFill { get; }
    /// <summary>Optional bounded hatch pattern.</summary>
    public OfficeChartHatchPattern? Hatch { get; }
    /// <summary>Hatch stroke colour.</summary>
    public OfficeColor? HatchColor { get; }
    /// <summary>Optional outline colour.</summary>
    public OfficeColor? OutlineColor { get; }
    /// <summary>Optional outline width in points.</summary>
    public double? OutlineWidth { get; }
    /// <summary>Outline visibility override; null inherits.</summary>
    public bool? ShowOutline { get; }
    /// <summary>Optional outline corner join.</summary>
    public OfficeStrokeLineJoin? OutlineJoin { get; }
}
