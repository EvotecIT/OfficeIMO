using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private bool TryStrokeDeviceHairline(IReadOnlyList<OfficeFlattenedPathContour> contours,
        OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit, IReadOnlyList<double>? pattern,
        double offset, OfficeTransform transform, Func<double, double, OfficeColor> paint) {
        OfficeTransform device = ScaleCoordinates(transform);
        if (!device.TryInvert(out OfficeTransform inverse)) return false;
        if (StrokeAdjustment && !(pattern is { Count: > 0 })) {
            var deviceContours = new List<OfficeFlattenedPathContour>(contours.Count);
            foreach (OfficeFlattenedPathContour contour in contours) {
                var points = new List<OfficePoint>(contour.Points.Count);
                foreach (OfficePoint point in contour.Points) points.Add(device.TransformPoint(point));
                deviceContours.Add(new OfficeFlattenedPathContour(points, contour.Closed, contour.UseStartLineCap, contour.UseEndLineCap));
            }
            // Snap in physical pixels: source-space width normalization cannot preserve the
            // normal under unequal axis scaling or shear.
            if (TryStrokeAdjustedContours(deviceContours, 1D, cap, join, miterLimit, null,
                OfficeTransform.Scale(1D / CoordinateScaleX, 1D / CoordinateScaleY), paint)) return true;
        }
        List<List<OfficePoint>> outlines = OfficeStrokeGeometry.CreateTransformedHairline(contours, cap, join, miterLimit,
            pattern, offset, device, inverse, Width, Height, _cancellationToken);
        var logical = new List<IReadOnlyList<OfficePoint>>(outlines.Count);
        foreach (List<OfficePoint> outline in outlines) {
            for (int index = 0; index < outline.Count; index++)
                outline[index] = new OfficePoint(outline[index].X / CoordinateScaleX, outline[index].Y / CoordinateScaleY);
            logical.Add(outline);
        }
        FillContourPaint(logical, OfficeFillRule.NonZero, paint);
        return true;
    }
}
