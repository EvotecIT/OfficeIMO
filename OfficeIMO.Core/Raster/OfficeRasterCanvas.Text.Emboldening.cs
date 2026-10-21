using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private void FillTextContours(IReadOnlyList<List<OfficePoint>> contours, OfficeColor color, double boldOffset = 0D) {
        var painted = TransformPositionedGlyphContours(contours);
        if (boldOffset == 0D) {
            if (_textInkObserver != null) { if (color.A != 0) _textInkObserver(painted); }
            else FillContours(painted, color, OfficeFillRule.NonZero);
            return;
        }
        // Synthetic bold is part of the glyph, so its offset precedes the glyph
        // transform too. The separate outline pen remains in destination units.
        FillTextContourUnion(painted, TransformPositionedGlyphContours(contours, boldOffset), color);
    }

    private void FillTextContourUnion(IReadOnlyList<List<OfficePoint>> contours, IReadOnlyList<List<OfficePoint>> shifted, OfficeColor color) {
        if (_textInkObserver != null) {
            if (color.A != 0) { _textInkObserver(contours); _textInkObserver(shifted); }
        } else FillContourPaint(contours, OfficeFillRule.NonZero, (_, _) => color, shifted);
    }
}
