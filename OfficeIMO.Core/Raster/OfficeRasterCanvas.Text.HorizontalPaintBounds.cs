using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Retain advances for the horizontal envelope contract, while sharing the
    // actual selected-font ink owner with placed raster segment bounds.
    internal (double Left, double Right) MeasureTextLineHorizontalPaintBounds(
        string text, double size, string? family, OfficeFontStyle style) =>
        MeasureTextLineHorizontalPaintBoundsWithFeatures(text, size, family, style, null);

    internal (double Left, double Right) MeasureTextLineHorizontalPaintBoundsWithFeatures(
        string text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
        var ink = MeasureTextLineInkBoundsWithFeatures(text, size, family, style, features);
        double advance = MeasureTextWithFeatures(text, Math.Max(1D, size), family, style, features);
        return ink.HasInk ? (Math.Min(0D, ink.Left), Math.Max(advance, ink.Right)) : (0D, advance);
    }
}
