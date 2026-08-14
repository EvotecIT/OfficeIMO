using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static bool CanvasClipIntersectsAnnotationRectangle(OfficeClipPath clipPath, double clipX, double clipBottomY, double clipHeight, double x1, double y1, double x2, double y2) {
            if (clipPath.Kind == OfficeClipPathKind.Empty) return false;
            if (clipPath.Kind == OfficeClipPathKind.Rectangle) return true;

            const int samplesPerAxis = 4;
            for (int yIndex = 0; yIndex <= samplesPerAxis; yIndex++) {
                double pageY = y1 + (y2 - y1) * yIndex / samplesPerAxis;
                double localY = clipHeight - (pageY - clipBottomY);
                for (int xIndex = 0; xIndex <= samplesPerAxis; xIndex++) {
                    double pageX = x1 + (x2 - x1) * xIndex / samplesPerAxis;
                    if (CanvasClipContainsPoint(clipPath, pageX - clipX, localY)) return true;
                }
            }

            if (clipPath.Kind == OfficeClipPathKind.RoundedRectangle) {
                return CanvasAnnotationContainsLocalPoint(clipX, clipBottomY, clipHeight, x1, y1, x2, y2, clipPath.Width / 2D, clipPath.Height / 2D);
            }
            if (clipPath.Kind == OfficeClipPathKind.Path) {
                foreach (List<OfficePoint> contour in FlattenCanvasClipContours(clipPath.Commands)) {
                    for (int index = 0; index < contour.Count; index++) {
                        OfficePoint point = contour[index];
                        if (CanvasAnnotationContainsLocalPoint(clipX, clipBottomY, clipHeight, x1, y1, x2, y2, point.X, point.Y)) return true;
                        OfficePoint next = contour[(index + 1) % contour.Count];
                        if (CanvasClipSegmentIntersectsAnnotation(clipX, clipBottomY, clipHeight, x1, y1, x2, y2, point, next)) return true;
                    }
                }
            }
            return false;
        }

        private static bool CanvasAnnotationContainsLocalPoint(double clipX, double clipBottomY, double clipHeight, double x1, double y1, double x2, double y2, double localX, double localY) {
            double pageX = clipX + localX;
            double pageY = clipBottomY + clipHeight - localY;
            return pageX >= x1 && pageX <= x2 && pageY >= y1 && pageY <= y2;
        }

        private static bool CanvasClipSegmentIntersectsAnnotation(double clipX, double clipBottomY, double clipHeight, double x1, double y1, double x2, double y2, OfficePoint a, OfficePoint b) {
            double left = x1 - clipX;
            double right = x2 - clipX;
            double top = clipHeight - (y2 - clipBottomY);
            double bottom = clipHeight - (y1 - clipBottomY);
            if (Math.Max(a.X, b.X) < left || Math.Min(a.X, b.X) > right || Math.Max(a.Y, b.Y) < top || Math.Min(a.Y, b.Y) > bottom) return false;
            return CanvasSegmentsIntersect(a, b, new OfficePoint(left, top), new OfficePoint(right, top))
                || CanvasSegmentsIntersect(a, b, new OfficePoint(right, top), new OfficePoint(right, bottom))
                || CanvasSegmentsIntersect(a, b, new OfficePoint(right, bottom), new OfficePoint(left, bottom))
                || CanvasSegmentsIntersect(a, b, new OfficePoint(left, bottom), new OfficePoint(left, top));
        }

        private static bool CanvasSegmentsIntersect(OfficePoint a, OfficePoint b, OfficePoint c, OfficePoint d) {
            const double tolerance = 0.001D;
            double abC = CanvasCross(a, b, c.X, c.Y);
            double abD = CanvasCross(a, b, d.X, d.Y);
            double cdA = CanvasCross(c, d, a.X, a.Y);
            double cdB = CanvasCross(c, d, b.X, b.Y);
            if ((abC > tolerance && abD < -tolerance || abC < -tolerance && abD > tolerance)
                && (cdA > tolerance && cdB < -tolerance || cdA < -tolerance && cdB > tolerance)) return true;
            return Math.Abs(abC) <= tolerance && CanvasPointOnSegment(c.X, c.Y, a, b, tolerance)
                || Math.Abs(abD) <= tolerance && CanvasPointOnSegment(d.X, d.Y, a, b, tolerance)
                || Math.Abs(cdA) <= tolerance && CanvasPointOnSegment(a.X, a.Y, c, d, tolerance)
                || Math.Abs(cdB) <= tolerance && CanvasPointOnSegment(b.X, b.Y, c, d, tolerance);
        }

        private static bool CanvasClipContainsPoint(OfficeClipPath clipPath, double x, double y) {
            const double tolerance = 0.001D;
            if (x < -tolerance || y < -tolerance || x > clipPath.Width + tolerance || y > clipPath.Height + tolerance) return false;
            if (clipPath.Kind == OfficeClipPathKind.RoundedRectangle) {
                double radius = clipPath.CornerRadius;
                if (radius <= tolerance || x >= radius && x <= clipPath.Width - radius || y >= radius && y <= clipPath.Height - radius) return true;
                double centerX = x < radius ? radius : clipPath.Width - radius;
                double centerY = y < radius ? radius : clipPath.Height - radius;
                double dx = x - centerX;
                double dy = y - centerY;
                return dx * dx + dy * dy <= radius * radius + tolerance;
            }
            if (clipPath.Kind != OfficeClipPathKind.Path) return true;

            List<List<OfficePoint>> contours = FlattenCanvasClipContours(clipPath.Commands);
            int winding = 0;
            bool odd = false;
            foreach (List<OfficePoint> contour in contours) {
                for (int index = 0, previous = contour.Count - 1; index < contour.Count; previous = index++) {
                    OfficePoint a = contour[previous];
                    OfficePoint b = contour[index];
                    if (CanvasPointOnSegment(x, y, a, b, tolerance)) return true;
                    if ((a.Y > y) != (b.Y > y) && x < (b.X - a.X) * (y - a.Y) / (b.Y - a.Y) + a.X) odd = !odd;
                    if (a.Y <= y) {
                        if (b.Y > y && CanvasCross(a, b, x, y) > 0D) winding++;
                    } else if (b.Y <= y && CanvasCross(a, b, x, y) < 0D) {
                        winding--;
                    }
                }
            }
            return clipPath.FillRule == OfficeFillRule.EvenOdd ? odd : winding != 0;
        }

        private static List<List<OfficePoint>> FlattenCanvasClipContours(IReadOnlyList<OfficePathCommand> commands) {
            const int curveSegments = 16;
            var contours = new List<List<OfficePoint>>();
            List<OfficePoint>? contour = null;
            OfficePoint current = default;
            foreach (OfficePathCommand command in commands) {
                if (command.Kind == OfficePathCommandKind.MoveTo) {
                    if (contour != null && contour.Count > 1) contours.Add(contour);
                    contour = new List<OfficePoint> { command.Point };
                    current = command.Point;
                } else if (contour != null && command.Kind == OfficePathCommandKind.LineTo) {
                    contour.Add(command.Point);
                    current = command.Point;
                } else if (contour != null && command.Kind == OfficePathCommandKind.QuadraticBezierTo) {
                    OfficePoint start = current;
                    for (int index = 1; index <= curveSegments; index++) {
                        double t = index / (double)curveSegments;
                        double inverse = 1D - t;
                        contour.Add(new OfficePoint(
                            inverse * inverse * start.X + 2D * inverse * t * command.ControlPoint1.X + t * t * command.Point.X,
                            inverse * inverse * start.Y + 2D * inverse * t * command.ControlPoint1.Y + t * t * command.Point.Y));
                    }
                    current = command.Point;
                } else if (contour != null && command.Kind == OfficePathCommandKind.CubicBezierTo) {
                    OfficePoint start = current;
                    for (int index = 1; index <= curveSegments; index++) {
                        double t = index / (double)curveSegments;
                        double inverse = 1D - t;
                        contour.Add(new OfficePoint(
                            inverse * inverse * inverse * start.X + 3D * inverse * inverse * t * command.ControlPoint1.X + 3D * inverse * t * t * command.ControlPoint2.X + t * t * t * command.Point.X,
                            inverse * inverse * inverse * start.Y + 3D * inverse * inverse * t * command.ControlPoint1.Y + 3D * inverse * t * t * command.ControlPoint2.Y + t * t * t * command.Point.Y));
                    }
                    current = command.Point;
                } else if (contour != null && command.Kind == OfficePathCommandKind.Close) {
                    if (contour.Count > 1) contours.Add(contour);
                    contour = null;
                }
            }
            if (contour != null && contour.Count > 1) contours.Add(contour);
            return contours;
        }

        private static bool CanvasPointOnSegment(double x, double y, OfficePoint a, OfficePoint b, double tolerance) {
            double cross = CanvasCross(a, b, x, y);
            if (Math.Abs(cross) > tolerance) return false;
            return x >= Math.Min(a.X, b.X) - tolerance && x <= Math.Max(a.X, b.X) + tolerance
                && y >= Math.Min(a.Y, b.Y) - tolerance && y <= Math.Max(a.Y, b.Y) + tolerance;
        }

        private static double CanvasCross(OfficePoint a, OfficePoint b, double x, double y) =>
            (b.X - a.X) * (y - a.Y) - (b.Y - a.Y) * (x - a.X);
    }
}
