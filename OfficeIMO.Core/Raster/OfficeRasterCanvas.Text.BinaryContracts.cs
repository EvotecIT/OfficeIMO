namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Published document assemblies call this exact fifteen-argument, five-value
    // signature. Keep buffer bounds separate from the extended ink-inspection result.
    internal (double Left, double Top, double Right, double Bottom, bool HasInk) MeasurePositionedTextBounds(
        string text, double x, double y, double width, double height, double size,
        OfficeFontInfo fontInfo, double advance, OfficeTextAlignment alignment,
        OfficeTextFeatureSettings features, string palette, double baselineSize,
        OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextDirection textDirection) {
        var bounds = MeasurePositionedTextBounds(text, x, y, width, height, size,
            fontInfo, advance, alignment, features, palette, baselineSize,
            underlineStyle, strikethroughStyle, textDirection, inkOnly: false);
        return (bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, bounds.HasInk);
    }
}
