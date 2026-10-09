using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Retain advances for the horizontal envelope contract, while sharing the
    // actual selected-font ink owner with placed raster segment bounds.
    internal (double Left, double Right) MeasureTextLineHorizontalPaintBounds(
        string text, double size, string? family, OfficeFontStyle style) {
        var ink = MeasureTextLineInkBounds(text, size, family, style);
        double advance = MeasureText(text, Math.Max(1D, size), family, style);
        return ink.HasInk ? (Math.Min(0D, ink.Left), Math.Max(advance, ink.Right)) : (0D, advance);
    }
}
