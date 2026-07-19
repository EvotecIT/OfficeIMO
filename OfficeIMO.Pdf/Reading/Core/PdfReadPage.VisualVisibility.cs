using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const double VisualGeometryEpsilon = 0.000001D;

    private static bool IsVisibleVisualPrimitive(
        PdfPageVisualPrimitive primitive,
        double pageWidth,
        double pageHeight) {
        if (!HasFinitePrimitiveGeometry(primitive) ||
            !IsFinite(pageWidth) ||
            !IsFinite(pageHeight) ||
            pageWidth <= 0D ||
            pageHeight <= 0D) {
            return false;
        }

        bool hasVisibleFill = primitive.Kind != PdfPageVisualPrimitiveKind.Line &&
            ((HasOrdinaryFill(primitive) && HasVisibleOpacity(primitive.FillOpacity)) ||
             HasVisibleTilingPattern(primitive.FillTilingPattern));
        bool hasVisibleStroke = primitive.StrokeWidth > 0D &&
            ((HasOrdinaryStroke(primitive) && HasVisibleOpacity(primitive.StrokeOpacity)) ||
             HasVisibleTilingPattern(primitive.StrokeTilingPattern));
        if (!hasVisibleFill && !hasVisibleStroke) {
            return false;
        }

        PdfPageClipPath pageClip = PdfPageClipPath.Rectangle(0D, 0D, pageWidth, pageHeight);
        PdfPageClipPath visibleClip = pageClip;
        if (primitive.ClipPath.HasValue) {
            PdfPageClipPath authoredClip = primitive.ClipPath.Value;
            if (!HasFiniteClipGeometry(authoredClip)) {
                return false;
            }

            visibleClip = PdfPageClipPath.ResolveActiveClip(pageClip, authoredClip);
        }

        if (!HasFiniteClipGeometry(visibleClip) ||
            visibleClip.Width <= 0D ||
            visibleClip.Height <= 0D) {
            return false;
        }

        VisualPath clipPath = VisualPath.FromClip(visibleClip);
        return (hasVisibleFill && VisualPath.FromFill(primitive).IntersectsFill(clipPath)) ||
            (hasVisibleStroke && VisualPath.FromStroke(primitive).StrokeIntersectsFill(clipPath, primitive.StrokeWidth / 2D));
    }

    private static bool HasFinitePrimitiveGeometry(PdfPageVisualPrimitive primitive) {
        if (!IsFinite(primitive.X) ||
            !IsFinite(primitive.Y) ||
            !IsFinite(primitive.Width) ||
            !IsFinite(primitive.Height) ||
            !IsFinite(primitive.X1) ||
            !IsFinite(primitive.Y1) ||
            !IsFinite(primitive.X2) ||
            !IsFinite(primitive.Y2) ||
            !IsFinite(primitive.StrokeWidth)) {
            return false;
        }

        for (int i = 0; i < primitive.PathCommands.Count; i++) {
            if (!HasFiniteCommand(primitive.PathCommands[i])) {
                return false;
            }
        }

        return true;
    }

    private static bool HasFiniteClipGeometry(PdfPageClipPath clip) {
        if (!IsFinite(clip.X) ||
            !IsFinite(clip.Y) ||
            !IsFinite(clip.Width) ||
            !IsFinite(clip.Height)) {
            return false;
        }

        for (int i = 0; i < clip.Commands.Count; i++) {
            if (!HasFiniteCommand(clip.Commands[i])) {
                return false;
            }
        }

        return true;
    }

    private static bool HasFiniteCommand(OfficePathCommand command) =>
        IsFinite(command.Point.X) &&
        IsFinite(command.Point.Y) &&
        IsFinite(command.ControlPoint1.X) &&
        IsFinite(command.ControlPoint1.Y) &&
        IsFinite(command.ControlPoint2.X) &&
        IsFinite(command.ControlPoint2.Y);

    private static bool HasOrdinaryFill(PdfPageVisualPrimitive primitive) =>
        HasVisibleColor(primitive.FillColor) ||
        HasVisibleGradient(primitive.FillGradient) ||
        HasVisibleGradient(primitive.FillRadialGradient);

    private static bool HasOrdinaryStroke(PdfPageVisualPrimitive primitive) =>
        HasVisibleColor(primitive.StrokeColor) ||
        HasVisibleGradient(primitive.StrokeGradient) ||
        HasVisibleGradient(primitive.StrokeRadialGradient);

    private static bool HasVisibleTilingPattern(PdfPageTilingPatternPaint? pattern) =>
        pattern != null &&
        IsFinite(pattern.Opacity) &&
        pattern.Opacity > 0D &&
        HasVisibleDrawingContent(pattern.Resource.Tile);

    private static bool HasVisibleDrawingContent(OfficeDrawing drawing) {
        for (int i = 0; i < drawing.Elements.Count; i++) {
            OfficeDrawingElement element = drawing.Elements[i];
            switch (element) {
                case OfficeDrawingShape shape:
                    if (HasVisibleShapePaint(shape.Shape)) {
                        return true;
                    }
                    break;
                case OfficeDrawingImage image:
                    if (IsFinite(image.Opacity) && image.Opacity > 0D) {
                        return true;
                    }
                    break;
                case OfficeDrawingImagePattern imagePattern:
                    if (IsFinite(imagePattern.Opacity) && imagePattern.Opacity > 0D) {
                        return true;
                    }
                    break;
                case OfficeDrawingText text:
                    if (!string.IsNullOrEmpty(text.Text) && (!text.Color.HasValue || text.Color.Value.A > 0)) {
                        return true;
                    }
                    break;
                case OfficeDrawingRichText richText:
                    if (!string.IsNullOrEmpty(richText.PlainText)) {
                        return true;
                    }
                    break;
                case OfficeDrawingGroup group:
                    if (HasVisibleDrawingContent(group.Drawing)) {
                        return true;
                    }
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (IsFinite(effectGroup.Opacity) &&
                        effectGroup.Opacity > 0D &&
                        HasVisibleSoftMask(effectGroup.SoftMask) &&
                        HasVisibleDrawingContent(effectGroup.Drawing)) {
                        return true;
                    }
                    break;
                case OfficeDrawingTilingPattern tilingPattern:
                    if (IsFinite(tilingPattern.Opacity) &&
                        tilingPattern.Opacity > 0D &&
                        HasVisibleDrawingContent(tilingPattern.Tile)) {
                        return true;
                    }
                    break;
            }
        }

        return false;
    }

    private static bool HasVisibleSoftMask(OfficeDrawingSoftMask? softMask) =>
        softMask == null ||
        softMask.BackdropColor.A > 0 ||
        HasVisibleDrawingContent(softMask.Drawing);

    private static bool HasVisibleShapePaint(OfficeShape shape) {
        bool fill = (HasVisibleColor(shape.FillColor) ||
                     HasVisibleGradient(shape.FillGradient) ||
                     HasVisibleGradient(shape.FillRadialGradient)) &&
            HasVisibleOpacity(shape.FillOpacity);
        bool stroke = shape.StrokeWidth > 0D &&
            (HasVisibleColor(shape.StrokeColor) ||
             HasVisibleGradient(shape.StrokeGradient) ||
             HasVisibleGradient(shape.StrokeRadialGradient)) &&
            HasVisibleOpacity(shape.StrokeOpacity);
        return fill || stroke;
    }

    private static bool HasVisibleColor(OfficeColor? color) =>
        color.HasValue && color.Value.A > 0;

    private static bool HasVisibleGradient(OfficeLinearGradient? gradient) =>
        gradient != null && gradient.Stops.Any(static stop => stop.Color.A > 0);

    private static bool HasVisibleGradient(OfficeRadialGradient? gradient) =>
        gradient != null && gradient.Stops.Any(static stop => stop.Color.A > 0);

    private static bool HasVisibleOpacity(double? opacity) =>
        !opacity.HasValue || (IsFinite(opacity.Value) && opacity.Value > 0D);

    private sealed class VisualPath {
        private readonly List<VisualContour> _contours;

        private VisualPath(List<VisualContour> contours, OfficeFillRule fillRule) {
            _contours = contours;
            FillRule = fillRule;
        }

        private OfficeFillRule FillRule { get; }

        public static VisualPath FromClip(PdfPageClipPath clip) {
            if (clip.IsRectangle) {
                return Rectangle(clip.X, clip.Y, clip.Width, clip.Height);
            }

            return FromCommands(clip.Commands, clip.FillRule);
        }

        public static VisualPath FromFill(PdfPageVisualPrimitive primitive) {
            return primitive.Kind == PdfPageVisualPrimitiveKind.Path
                ? FromCommands(primitive.PathCommands, primitive.FillRule)
                : Rectangle(primitive.X, primitive.Y, primitive.Width, primitive.Height);
        }

        public static VisualPath FromStroke(PdfPageVisualPrimitive primitive) {
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                return new VisualPath(
                    new List<VisualContour> {
                        new VisualContour(
                            new List<OfficePoint> {
                                new OfficePoint(primitive.X1, primitive.Y1),
                                new OfficePoint(primitive.X2, primitive.Y2)
                            },
                            closed: false)
                    },
                    OfficeFillRule.NonZero);
            }

            return primitive.Kind == PdfPageVisualPrimitiveKind.Path
                ? FromCommands(primitive.PathCommands, primitive.FillRule)
                : Rectangle(primitive.X, primitive.Y, primitive.Width, primitive.Height);
        }

        public bool IntersectsFill(VisualPath other) {
            if (_contours.Count == 0 || other._contours.Count == 0) {
                return false;
            }

            if (BoundariesIntersect(other)) {
                return true;
            }

            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                List<OfficePoint> points = _contours[contourIndex].Points;
                for (int pointIndex = 0; pointIndex < points.Count; pointIndex++) {
                    if (other.Contains(points[pointIndex])) {
                        return true;
                    }
                }
            }

            for (int contourIndex = 0; contourIndex < other._contours.Count; contourIndex++) {
                List<OfficePoint> points = other._contours[contourIndex].Points;
                for (int pointIndex = 0; pointIndex < points.Count; pointIndex++) {
                    if (Contains(points[pointIndex])) {
                        return true;
                    }
                }
            }

            return false;
        }

        public bool StrokeIntersectsFill(VisualPath fill, double strokeHalfWidth) {
            if (_contours.Count == 0 ||
                fill._contours.Count == 0 ||
                !IsFinite(strokeHalfWidth) ||
                strokeHalfWidth <= 0D) {
                return false;
            }

            double maximumDistanceSquared = strokeHalfWidth * strokeHalfWidth;
            if (!IsFinite(maximumDistanceSquared)) {
                return false;
            }

            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                int segmentCount = contour.SegmentCount(closeForFill: false);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    contour.GetSegment(segmentIndex, closeForFill: false, out OfficePoint start, out OfficePoint end);
                    OfficePoint midpoint = new OfficePoint((start.X + end.X) / 2D, (start.Y + end.Y) / 2D);
                    if (fill.Contains(start) || fill.Contains(end) || fill.Contains(midpoint)) {
                        return true;
                    }

                    for (int fillContourIndex = 0; fillContourIndex < fill._contours.Count; fillContourIndex++) {
                        VisualContour fillContour = fill._contours[fillContourIndex];
                        int fillSegmentCount = fillContour.SegmentCount(closeForFill: true);
                        for (int fillSegmentIndex = 0; fillSegmentIndex < fillSegmentCount; fillSegmentIndex++) {
                            fillContour.GetSegment(fillSegmentIndex, closeForFill: true, out OfficePoint fillStart, out OfficePoint fillEnd);
                            if (SegmentDistanceSquared(start, end, fillStart, fillEnd) <= maximumDistanceSquared + VisualGeometryEpsilon) {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private bool BoundariesIntersect(VisualPath other) {
            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                int segmentCount = contour.SegmentCount(closeForFill: true);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    contour.GetSegment(segmentIndex, closeForFill: true, out OfficePoint start, out OfficePoint end);
                    for (int otherContourIndex = 0; otherContourIndex < other._contours.Count; otherContourIndex++) {
                        VisualContour otherContour = other._contours[otherContourIndex];
                        int otherSegmentCount = otherContour.SegmentCount(closeForFill: true);
                        for (int otherSegmentIndex = 0; otherSegmentIndex < otherSegmentCount; otherSegmentIndex++) {
                            otherContour.GetSegment(otherSegmentIndex, closeForFill: true, out OfficePoint otherStart, out OfficePoint otherEnd);
                            if (SegmentsIntersect(start, end, otherStart, otherEnd)) {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private bool Contains(OfficePoint point) {
            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                int segmentCount = contour.SegmentCount(closeForFill: true);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    contour.GetSegment(segmentIndex, closeForFill: true, out OfficePoint start, out OfficePoint end);
                    if (PointOnSegment(point, start, end)) {
                        return true;
                    }
                }
            }

            if (FillRule == OfficeFillRule.EvenOdd) {
                bool inside = false;
                for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                    VisualContour contour = _contours[contourIndex];
                    int segmentCount = contour.SegmentCount(closeForFill: true);
                    for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                        contour.GetSegment(segmentIndex, closeForFill: true, out OfficePoint start, out OfficePoint end);
                        if ((start.Y > point.Y) != (end.Y > point.Y) &&
                            point.X < ((end.X - start.X) * (point.Y - start.Y) / (end.Y - start.Y)) + start.X) {
                            inside = !inside;
                        }
                    }
                }

                return inside;
            }

            int winding = 0;
            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                int segmentCount = contour.SegmentCount(closeForFill: true);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    contour.GetSegment(segmentIndex, closeForFill: true, out OfficePoint start, out OfficePoint end);
                    double cross = Cross(start, end, point);
                    if (start.Y <= point.Y) {
                        if (end.Y > point.Y && cross > VisualGeometryEpsilon) {
                            winding++;
                        }
                    } else if (end.Y <= point.Y && cross < -VisualGeometryEpsilon) {
                        winding--;
                    }
                }
            }

            return winding != 0;
        }

        private static VisualPath Rectangle(double x, double y, double width, double height) {
            var points = new List<OfficePoint> {
                new OfficePoint(x, y),
                new OfficePoint(x + width, y),
                new OfficePoint(x + width, y + height),
                new OfficePoint(x, y + height)
            };
            return new VisualPath(
                new List<VisualContour> { new VisualContour(points, closed: true) },
                OfficeFillRule.NonZero);
        }

        private static VisualPath FromCommands(IReadOnlyList<OfficePathCommand> commands, OfficeFillRule fillRule) {
            var contours = new List<VisualContour>();
            List<OfficePoint>? current = null;
            OfficePoint currentPoint = default;
            for (int i = 0; i < commands.Count; i++) {
                OfficePathCommand command = commands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        AddContour(contours, current, closed: false);
                        currentPoint = command.Point;
                        current = new List<OfficePoint> { currentPoint };
                        break;
                    case OfficePathCommandKind.LineTo:
                        if (current == null) {
                            current = new List<OfficePoint> { currentPoint };
                        }
                        currentPoint = command.Point;
                        current.Add(currentPoint);
                        break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        if (current == null) {
                            current = new List<OfficePoint> { currentPoint };
                        }
                        current.AddRange(OfficeGeometry.CreateQuadraticBezierPoints(currentPoint, command.ControlPoint1, command.Point, 24));
                        currentPoint = command.Point;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        if (current == null) {
                            current = new List<OfficePoint> { currentPoint };
                        }
                        current.AddRange(OfficeGeometry.CreateCubicBezierPoints(currentPoint, command.ControlPoint1, command.ControlPoint2, command.Point, 24));
                        currentPoint = command.Point;
                        break;
                    case OfficePathCommandKind.Close:
                        AddContour(contours, current, closed: true);
                        current = null;
                        break;
                }
            }

            AddContour(contours, current, closed: false);
            return new VisualPath(contours, fillRule);
        }

        private static void AddContour(List<VisualContour> contours, List<OfficePoint>? points, bool closed) {
            if (points == null || points.Count < 2) {
                return;
            }

            if (points.Count > 1 && PointsEqual(points[0], points[points.Count - 1])) {
                points.RemoveAt(points.Count - 1);
                closed = true;
            }

            if (points.Count >= 2) {
                contours.Add(new VisualContour(points, closed));
            }
        }
    }

    private sealed class VisualContour {
        public VisualContour(List<OfficePoint> points, bool closed) {
            Points = points;
            Closed = closed;
        }

        public List<OfficePoint> Points { get; }
        private bool Closed { get; }

        public int SegmentCount(bool closeForFill) =>
            Points.Count < 2 ? 0 : Points.Count - 1 + ((Closed || closeForFill) ? 1 : 0);

        public void GetSegment(int index, bool closeForFill, out OfficePoint start, out OfficePoint end) {
            start = Points[index];
            end = index + 1 < Points.Count
                ? Points[index + 1]
                : Points[0];
        }
    }

    private static bool SegmentsIntersect(OfficePoint firstStart, OfficePoint firstEnd, OfficePoint secondStart, OfficePoint secondEnd) {
        double firstCrossStart = Cross(firstStart, firstEnd, secondStart);
        double firstCrossEnd = Cross(firstStart, firstEnd, secondEnd);
        double secondCrossStart = Cross(secondStart, secondEnd, firstStart);
        double secondCrossEnd = Cross(secondStart, secondEnd, firstEnd);
        if (((firstCrossStart > VisualGeometryEpsilon && firstCrossEnd < -VisualGeometryEpsilon) ||
             (firstCrossStart < -VisualGeometryEpsilon && firstCrossEnd > VisualGeometryEpsilon)) &&
            ((secondCrossStart > VisualGeometryEpsilon && secondCrossEnd < -VisualGeometryEpsilon) ||
             (secondCrossStart < -VisualGeometryEpsilon && secondCrossEnd > VisualGeometryEpsilon))) {
            return true;
        }

        return (Math.Abs(firstCrossStart) <= VisualGeometryEpsilon && PointOnSegment(secondStart, firstStart, firstEnd)) ||
            (Math.Abs(firstCrossEnd) <= VisualGeometryEpsilon && PointOnSegment(secondEnd, firstStart, firstEnd)) ||
            (Math.Abs(secondCrossStart) <= VisualGeometryEpsilon && PointOnSegment(firstStart, secondStart, secondEnd)) ||
            (Math.Abs(secondCrossEnd) <= VisualGeometryEpsilon && PointOnSegment(firstEnd, secondStart, secondEnd));
    }

    private static bool PointOnSegment(OfficePoint point, OfficePoint start, OfficePoint end) {
        if (Math.Abs(Cross(start, end, point)) > VisualGeometryEpsilon) {
            return false;
        }

        return point.X >= Math.Min(start.X, end.X) - VisualGeometryEpsilon &&
            point.X <= Math.Max(start.X, end.X) + VisualGeometryEpsilon &&
            point.Y >= Math.Min(start.Y, end.Y) - VisualGeometryEpsilon &&
            point.Y <= Math.Max(start.Y, end.Y) + VisualGeometryEpsilon;
    }

    private static double SegmentDistanceSquared(
        OfficePoint firstStart,
        OfficePoint firstEnd,
        OfficePoint secondStart,
        OfficePoint secondEnd) {
        if (SegmentsIntersect(firstStart, firstEnd, secondStart, secondEnd)) {
            return 0D;
        }

        return Math.Min(
            Math.Min(PointSegmentDistanceSquared(firstStart, secondStart, secondEnd), PointSegmentDistanceSquared(firstEnd, secondStart, secondEnd)),
            Math.Min(PointSegmentDistanceSquared(secondStart, firstStart, firstEnd), PointSegmentDistanceSquared(secondEnd, firstStart, firstEnd)));
    }

    private static double PointSegmentDistanceSquared(OfficePoint point, OfficePoint start, OfficePoint end) {
        double deltaX = end.X - start.X;
        double deltaY = end.Y - start.Y;
        double lengthSquared = (deltaX * deltaX) + (deltaY * deltaY);
        if (lengthSquared <= VisualGeometryEpsilon) {
            double pointDeltaX = point.X - start.X;
            double pointDeltaY = point.Y - start.Y;
            return (pointDeltaX * pointDeltaX) + (pointDeltaY * pointDeltaY);
        }

        double projection = (((point.X - start.X) * deltaX) + ((point.Y - start.Y) * deltaY)) / lengthSquared;
        projection = Math.Max(0D, Math.Min(1D, projection));
        double closestX = start.X + (projection * deltaX);
        double closestY = start.Y + (projection * deltaY);
        double distanceX = point.X - closestX;
        double distanceY = point.Y - closestY;
        return (distanceX * distanceX) + (distanceY * distanceY);
    }

    private static double Cross(OfficePoint start, OfficePoint end, OfficePoint point) =>
        ((end.X - start.X) * (point.Y - start.Y)) -
        ((end.Y - start.Y) * (point.X - start.X));

    private static bool PointsEqual(OfficePoint left, OfficePoint right) =>
        Math.Abs(left.X - right.X) <= VisualGeometryEpsilon &&
        Math.Abs(left.Y - right.Y) <= VisualGeometryEpsilon;
}
