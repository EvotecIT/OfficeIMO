using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    private static void DrawRasterTabLineLeader(OfficeRasterCanvas canvas, OfficeTextTabLineLeaderPaint paint,
        double x, double baseline, double rotationDegrees, double centerX, double centerY, bool flipHorizontal, bool flipVertical) {
        var contours = new List<IReadOnlyList<OfficePoint>>(paint.Contours.Count);
        double radians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        foreach (var contour in paint.Contours) {
            canvas.CancellationToken.ThrowIfCancellationRequested();
            var transformed = new List<OfficePoint>(contour.Count);
            foreach (OfficePoint point in contour) transformed.Add(TransformTextRectanglePoint(
                new OfficePoint(x + point.X, baseline + point.Y), radians, centerX, centerY, flipHorizontal, flipVertical));
            contours.Add(transformed);
        }
        // One non-zero union prevents overlapping wave outline pieces from blending
        // translucent color more than once.
        canvas.FillContourPaint(contours, OfficeFillRule.NonZero, (_, _) => paint.Color);
    }

    private static StringBuilder AppendSvgTabLineLeader(StringBuilder builder, OfficeTextTabLineLeaderPaint paint,
        double x, double baseline, double rotationDegrees, double centerX, double centerY) {
        builder.Append("<path").AppendPaintAttribute("fill", paint.Color).Append(" fill-rule=\"nonzero\"");
        if (Math.Abs(rotationDegrees) > .000001D) builder.AppendRotateTransformAttribute(rotationDegrees, centerX, centerY);
        builder.Append(" d=\"");
        foreach (var contour in paint.Contours) {
            for (int i = 0; i < contour.Count; i++) {
                builder.Append(i == 0 ? 'M' : 'L')
                    .Append((x + contour[i].X).ToString("G17", CultureInfo.InvariantCulture)).Append(' ')
                    .Append((baseline + contour[i].Y).ToString("G17", CultureInfo.InvariantCulture));
            }
            builder.Append('Z');
        }
        return builder.Append("\"/>");
    }
}
