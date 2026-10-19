using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeStrokeGeometry {
    // Dash distances belong to the source path; the one-pixel outline belongs to the device.
    internal static List<List<OfficePoint>> CreateTransformedHairline(IReadOnlyList<OfficeFlattenedPathContour> contours,
        OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit, IReadOnlyList<double>? dashPattern,
        double dashOffset, OfficeTransform deviceTransform, OfficeTransform inverse, int width, int height,
        CancellationToken cancellationToken) {
        if (!IsFinite(miterLimit) || miterLimit < 1D) miterLimit = 4D;
        double density = LocalFlatteningDensity(deviceTransform, 1D);
        IReadOnlyList<double> pattern = NormalizeRasterDashPattern(NormalizeDashPattern(dashPattern), .25D / density);
        double cycle = 0D;
        foreach (double length in pattern) cycle = cycle > double.MaxValue - length ? double.MaxValue : cycle + length;
        var bounds = inverse.TransformRectangleBounds(0D, 0D, width, height);
        double localPadding = LocalFlatteningDensity(inverse, 1D);
        var outlines = new List<List<OfficePoint>>();
        foreach (OfficeFlattenedPathContour contour in contours) {
            cancellationToken.ThrowIfCancellationRequested();
            var points = new List<OfficePoint>(contour.Points);
            if (contour.Closed && points.Count > 1 && !SameStrokePoint(points[0], points[points.Count - 1])) points.Add(points[0]);
            foreach (List<OfficePoint> run in CollectStrokeRuns(points, localPadding, miterLimit, pattern, cycle, dashOffset,
                         false, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int index = 0; index < run.Count; index++) run[index] = deviceTransform.TransformPoint(run[index]);
                OfficeRasterStroker.AppendOutline(run, 1D, cap, join, miterLimit, 1D, outlines);
            }
        }
        return outlines;
    }
}
