using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static bool TryClipCanvasAnnotationRectangle(OfficeClipPath clipPath, double clipX, double clipBottomY, double clipHeight, ref double x1, ref double y1, ref double x2, ref double y2) {
            if (clipPath.Kind == OfficeClipPathKind.Empty) return false;
            if (clipPath.Kind == OfficeClipPathKind.Rectangle) return true;

            const int gridSize = 24;
            double localLeft = x1 - clipX;
            double localRight = x2 - clipX;
            double localTop = clipHeight - (y2 - clipBottomY);
            double localBottom = clipHeight - (y1 - clipBottomY);
            double cellWidth = (localRight - localLeft) / gridSize;
            double cellHeight = (localBottom - localTop) / gridSize;
            List<List<OfficePoint>>? contours = clipPath.Kind == OfficeClipPathKind.Path
                ? FlattenCanvasClipContours(clipPath.Commands)
                : null;
            var safe = new bool[gridSize, gridSize];
            for (int row = 0; row < gridSize; row++) {
                double top = localTop + row * cellHeight;
                double bottom = top + cellHeight;
                for (int column = 0; column < gridSize; column++) {
                    double left = localLeft + column * cellWidth;
                    double right = left + cellWidth;
                    safe[row, column] = CanvasClipContainsCell(clipPath, contours, left, top, right, bottom);
                }
            }

            int bestTop = 0;
            int bestBottom = 0;
            int bestLeft = 0;
            int bestRight = 0;
            int bestArea = 0;
            var availableColumns = new bool[gridSize];
            for (int topRow = 0; topRow < gridSize; topRow++) {
                for (int column = 0; column < gridSize; column++) availableColumns[column] = true;
                for (int bottomRow = topRow; bottomRow < gridSize; bottomRow++) {
                    for (int column = 0; column < gridSize; column++) availableColumns[column] &= safe[bottomRow, column];
                    int runStart = -1;
                    for (int column = 0; column <= gridSize; column++) {
                        if (column < gridSize && availableColumns[column]) {
                            if (runStart < 0) runStart = column;
                            continue;
                        }
                        if (runStart < 0) continue;
                        int area = (bottomRow - topRow + 1) * (column - runStart);
                        if (area > bestArea) {
                            bestArea = area;
                            bestTop = topRow;
                            bestBottom = bottomRow + 1;
                            bestLeft = runStart;
                            bestRight = column;
                        }
                        runStart = -1;
                    }
                }
            }

            if (bestArea == 0) return false;
            double clippedLocalLeft = localLeft + bestLeft * cellWidth;
            double clippedLocalRight = localLeft + bestRight * cellWidth;
            double clippedLocalTop = localTop + bestTop * cellHeight;
            double clippedLocalBottom = localTop + bestBottom * cellHeight;
            x1 = clipX + clippedLocalLeft;
            x2 = clipX + clippedLocalRight;
            y1 = clipBottomY + clipHeight - clippedLocalBottom;
            y2 = clipBottomY + clipHeight - clippedLocalTop;
            return x2 > x1 && y2 > y1;
        }

        private static bool CanvasClipContainsCell(OfficeClipPath clipPath, List<List<OfficePoint>>? contours, double left, double top, double right, double bottom) {
            double middleX = (left + right) / 2D;
            double middleY = (top + bottom) / 2D;
            return CanvasClipContainsPoint(clipPath, contours, left, top)
                && CanvasClipContainsPoint(clipPath, contours, middleX, top)
                && CanvasClipContainsPoint(clipPath, contours, right, top)
                && CanvasClipContainsPoint(clipPath, contours, left, middleY)
                && CanvasClipContainsPoint(clipPath, contours, middleX, middleY)
                && CanvasClipContainsPoint(clipPath, contours, right, middleY)
                && CanvasClipContainsPoint(clipPath, contours, left, bottom)
                && CanvasClipContainsPoint(clipPath, contours, middleX, bottom)
                && CanvasClipContainsPoint(clipPath, contours, right, bottom)
                && (contours == null || !CanvasClipBoundaryIntersectsCell(contours, left, top, right, bottom));
        }

        private static bool CanvasClipContainsPoint(OfficeClipPath clipPath, List<List<OfficePoint>>? contours, double x, double y) {
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

            int winding = 0;
            bool odd = false;
            foreach (List<OfficePoint> contour in contours!) {
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

        private static bool CanvasClipBoundaryIntersectsCell(List<List<OfficePoint>> contours, double left, double top, double right, double bottom) {
            var topLeft = new OfficePoint(left, top);
            var topRight = new OfficePoint(right, top);
            var bottomRight = new OfficePoint(right, bottom);
            var bottomLeft = new OfficePoint(left, bottom);
            foreach (List<OfficePoint> contour in contours) {
                for (int index = 0; index < contour.Count; index++) {
                    OfficePoint a = contour[index];
                    OfficePoint b = contour[(index + 1) % contour.Count];
                    if (a.X > left && a.X < right && a.Y > top && a.Y < bottom) return true;
                    if (CanvasSegmentsIntersect(a, b, topLeft, topRight)
                        || CanvasSegmentsIntersect(a, b, topRight, bottomRight)
                        || CanvasSegmentsIntersect(a, b, bottomRight, bottomLeft)
                        || CanvasSegmentsIntersect(a, b, bottomLeft, topLeft)) return true;
                }
            }
            return false;
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
