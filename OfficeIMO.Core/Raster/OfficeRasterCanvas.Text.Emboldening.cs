using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private void FillTextContours(IReadOnlyList<List<OfficePoint>> contours, OfficeColor color, double boldOffset = 0D) {
        if (boldOffset == 0D) {
            FillContours(contours, color, OfficeFillRule.NonZero);
            return;
        }
        var shifted = new List<List<OfficePoint>>(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            var copy = new List<OfficePoint>(contour.Count);
            foreach (OfficePoint point in contour) copy.Add(new OfficePoint(point.X + boldOffset, point.Y));
            shifted.Add(copy);
        }
        FillContourPaint(contours, OfficeFillRule.NonZero, (_, _) => color, shifted);
    }

    private void FillTextContourUnion(IReadOnlyList<List<OfficePoint>> contours, IReadOnlyList<List<OfficePoint>> shifted, OfficeColor color) =>
        FillContourPaint(contours, OfficeFillRule.NonZero, (_, _) => color, shifted);
}
