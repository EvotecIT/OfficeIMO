using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const double VisualGeometryEpsilon = 0.000001D;

    private sealed class VisualGeometryBudget {
        private const int MaximumOperations = 250000;
        private const int MaximumFlattenedPoints = 32768;
        private int _operations;
        private int _flattenedPoints;

        public bool Exceeded { get; private set; }

        public bool TryUseOperation() {
            if (Exceeded || _operations >= MaximumOperations) {
                Exceeded = true;
                return false;
            }

            _operations++;
            return true;
        }

        public bool TryAddPoints(int count) {
            if (count < 0 ||
                Exceeded ||
                count > MaximumFlattenedPoints - _flattenedPoints) {
                Exceeded = true;
                return false;
            }

            _flattenedPoints += count;
            return true;
        }

        public void Exhaust() => Exceeded = true;
    }

    private sealed partial class VisualPath {
        private readonly List<VisualContour> _contours;

        private VisualPath(List<VisualContour> contours, OfficeFillRule fillRule) {
            _contours = contours;
            FillRule = fillRule;
            Bounds = VisualBounds.FromContours(contours);
        }

        private OfficeFillRule FillRule { get; }
        private VisualBounds Bounds { get; }

        public VisualBounds GetBounds(double expansion = 0D) =>
            expansion > 0D ? Bounds.Expand(expansion) : Bounds;

        public static bool TryGetCommonBounds(
            IReadOnlyList<VisualPath> paths,
            out VisualBounds common) {
            common = default;
            if (paths.Count == 0) {
                return false;
            }

            common = paths[0].Bounds;
            if (!common.HasPositiveArea) {
                return false;
            }

            for (int i = 1; i < paths.Count; i++) {
                if (!common.TryIntersectPositive(paths[i].Bounds, out common)) {
                    return false;
                }
            }

            return true;
        }

        public static VisualPath? FromClip(
            PdfPageClipPath clip,
            VisualGeometryBudget budget) {
            if (clip.IsRectangle) {
                return Rectangle(
                    clip.X,
                    clip.Y,
                    clip.Width,
                    clip.Height,
                    OfficeTransform.Identity,
                    budget);
            }

            return FromCommands(
                clip.Commands,
                clip.FillRule,
                OfficeTransform.Identity,
                budget);
        }

        public static VisualPath? FromFill(
            PdfPageVisualPrimitive primitive,
            VisualGeometryBudget budget) =>
            primitive.Kind == PdfPageVisualPrimitiveKind.Path
                ? FromCommands(
                    primitive.PathCommands,
                    primitive.FillRule,
                    OfficeTransform.Identity,
                    budget)
                : Rectangle(
                    primitive.X,
                    primitive.Y,
                    primitive.Width,
                    primitive.Height,
                    OfficeTransform.Identity,
                    budget);

        public static VisualPath? FromStroke(
            PdfPageVisualPrimitive primitive,
            VisualGeometryBudget budget) {
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                return FromPoints(
                    new[] {
                        new OfficePoint(primitive.X1, primitive.Y1),
                        new OfficePoint(primitive.X2, primitive.Y2)
                    },
                    closed: false,
                    OfficeFillRule.NonZero,
                    OfficeTransform.Identity,
                    budget);
            }

            return primitive.Kind == PdfPageVisualPrimitiveKind.Path
                ? FromCommands(
                    primitive.PathCommands,
                    primitive.FillRule,
                    OfficeTransform.Identity,
                    budget)
                : Rectangle(
                    primitive.X,
                    primitive.Y,
                    primitive.Width,
                    primitive.Height,
                    OfficeTransform.Identity,
                    budget);
        }

        public static VisualPath? FromOfficeClip(
            OfficeClipPath clip,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            switch (clip.Kind) {
                case OfficeClipPathKind.Rectangle:
                    return Rectangle(0D, 0D, clip.Width, clip.Height, transform, budget);
                case OfficeClipPathKind.RoundedRectangle:
                    return RoundedRectangle(
                        0D,
                        0D,
                        clip.Width,
                        clip.Height,
                        clip.CornerRadius,
                        transform,
                        budget);
                case OfficeClipPathKind.Path:
                    return FromCommands(clip.Commands, clip.FillRule, transform, budget);
                default:
                    return null;
            }
        }

        public static VisualPath? FromShape(
            OfficeShape shape,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            switch (shape.Kind) {
                case OfficeShapeKind.Rectangle:
                    return Rectangle(0D, 0D, shape.Width, shape.Height, transform, budget);
                case OfficeShapeKind.RoundedRectangle:
                    return RoundedRectangle(
                        0D,
                        0D,
                        shape.Width,
                        shape.Height,
                        shape.CornerRadius,
                        transform,
                        budget);
                case OfficeShapeKind.Ellipse:
                    return Ellipse(0D, 0D, shape.Width, shape.Height, transform, budget);
                case OfficeShapeKind.Polygon:
                    return FromPoints(
                        shape.Points,
                        closed: true,
                        shape.FillRule,
                        transform,
                        budget);
                case OfficeShapeKind.Path:
                    return FromCommands(shape.PathCommands, shape.FillRule, transform, budget);
                case OfficeShapeKind.Line:
                    return FromPoints(
                        shape.Points,
                        closed: false,
                        OfficeFillRule.NonZero,
                        transform,
                        budget);
                default:
                    return null;
            }
        }

        public static VisualPath? Rectangle(
            double x,
            double y,
            double width,
            double height,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            var points = new[] {
                new OfficePoint(x, y),
                new OfficePoint(x + width, y),
                new OfficePoint(x + width, y + height),
                new OfficePoint(x, y + height)
            };
            return FromPoints(
                points,
                closed: true,
                OfficeFillRule.NonZero,
                transform,
                budget);
        }

        public static double GetMaximumScale(OfficeTransform transform) {
            double horizontal = Math.Sqrt(
                (transform.M11 * transform.M11) +
                (transform.M12 * transform.M12));
            double vertical = Math.Sqrt(
                (transform.M21 * transform.M21) +
                (transform.M22 * transform.M22));
            double scale = Math.Max(horizontal, vertical);
            return IsFinite(scale) ? scale : 0D;
        }

        public bool IntersectsFill(
            VisualPath other,
            VisualGeometryBudget budget) =>
            HasPositiveAreaIntersection(new[] { this, other }, budget);

        public bool IntersectsFills(
            IReadOnlyList<VisualPath> fills,
            VisualGeometryBudget budget) {
            var paths = new List<VisualPath>(fills.Count + 1) { this };
            for (int i = 0; i < fills.Count; i++) {
                paths.Add(fills[i]);
            }

            return HasPositiveAreaIntersection(paths, budget);
        }

        public bool StrokeIntersectsFill(
            VisualPath fill,
            double strokeHalfWidth,
            VisualGeometryBudget budget) =>
            StrokeIntersectsFills(new[] { fill }, strokeHalfWidth, budget);

        public bool StrokeIntersectsFills(
            IReadOnlyList<VisualPath> fills,
            double strokeHalfWidth,
            VisualGeometryBudget budget) {
            if (_contours.Count == 0 ||
                fills.Count == 0 ||
                !IsFinite(strokeHalfWidth) ||
                strokeHalfWidth <= 0D ||
                !Bounds.Expand(strokeHalfWidth).HasPositiveArea) {
                return false;
            }

            for (int i = 0; i < fills.Count; i++) {
                if (!StrokeIntersectsSingleFill(fills[i], strokeHalfWidth, budget)) {
                    return budget.Exceeded;
                }
            }

            if (fills.Count == 1) {
                return true;
            }

            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                int segmentCount = contour.SegmentCount(closeForFill: false);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    contour.GetSegment(segmentIndex, closeForFill: false, out OfficePoint start, out OfficePoint end);
                    if (TryStrokeInteriorSamples(start, end, strokeHalfWidth, fills, budget)) {
                        return true;
                    }

                    if (budget.Exceeded) {
                        return true;
                    }
                }
            }

            return true;
        }

        private static VisualPath? RoundedRectangle(
            double x,
            double y,
            double width,
            double height,
            double radius,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            radius = Math.Max(0D, Math.Min(radius, Math.Min(width, height) / 2D));
            if (radius <= VisualGeometryEpsilon) {
                return Rectangle(x, y, width, height, transform, budget);
            }

            const int pointsPerCorner = 6;
            var points = new List<OfficePoint>(pointsPerCorner * 4);
            AddArc(points, x + radius, y + radius, radius, Math.PI, Math.PI * 1.5D, pointsPerCorner);
            AddArc(points, x + width - radius, y + radius, radius, Math.PI * 1.5D, Math.PI * 2D, pointsPerCorner);
            AddArc(points, x + width - radius, y + height - radius, radius, 0D, Math.PI * 0.5D, pointsPerCorner);
            AddArc(points, x + radius, y + height - radius, radius, Math.PI * 0.5D, Math.PI, pointsPerCorner);
            return FromPoints(points, closed: true, OfficeFillRule.NonZero, transform, budget);
        }

        private static VisualPath? Ellipse(
            double x,
            double y,
            double width,
            double height,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            const int pointCount = 32;
            var points = new List<OfficePoint>(pointCount);
            double centerX = x + (width / 2D);
            double centerY = y + (height / 2D);
            double radiusX = width / 2D;
            double radiusY = height / 2D;
            for (int i = 0; i < pointCount; i++) {
                double angle = Math.PI * 2D * i / pointCount;
                points.Add(new OfficePoint(
                    centerX + (Math.Cos(angle) * radiusX),
                    centerY + (Math.Sin(angle) * radiusY)));
            }

            return FromPoints(points, closed: true, OfficeFillRule.NonZero, transform, budget);
        }

        private static void AddArc(
            List<OfficePoint> points,
            double centerX,
            double centerY,
            double radius,
            double startAngle,
            double endAngle,
            int count) {
            for (int i = 0; i < count; i++) {
                double fraction = i / (double)(count - 1);
                double angle = startAngle + ((endAngle - startAngle) * fraction);
                points.Add(new OfficePoint(
                    centerX + (Math.Cos(angle) * radius),
                    centerY + (Math.Sin(angle) * radius)));
            }
        }

        private static VisualPath? FromPoints(
            IReadOnlyList<OfficePoint> source,
            bool closed,
            OfficeFillRule fillRule,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            if (source.Count < 2 || !budget.TryAddPoints(source.Count)) {
                return null;
            }

            var points = new List<OfficePoint>(source.Count);
            for (int i = 0; i < source.Count; i++) {
                OfficePoint point = transform.TransformPoint(source[i]);
                if (!IsFinite(point.X) || !IsFinite(point.Y)) {
                    budget.Exhaust();
                    return null;
                }
                points.Add(point);
            }

            if (points.Count > 1 && PointsEqual(points[0], points[points.Count - 1])) {
                points.RemoveAt(points.Count - 1);
                closed = true;
            }

            return points.Count < 2
                ? null
                : new VisualPath(
                    new List<VisualContour> { new VisualContour(points, closed) },
                    fillRule);
        }

        private static VisualPath? FromCommands(
            IReadOnlyList<OfficePathCommand> commands,
            OfficeFillRule fillRule,
            OfficeTransform transform,
            VisualGeometryBudget budget) {
            var contours = new List<VisualContour>();
            List<OfficePoint>? current = null;
            OfficePoint currentPoint = default;
            OfficePoint subpathStart = default;
            bool hasCurrentPoint = false;
            bool hasSubpathStart = false;
            for (int i = 0; i < commands.Count; i++) {
                OfficePathCommand command = commands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        AddContour(contours, current, closed: false);
                        if (!TryTransformAndReserve(command.Point, transform, budget, out currentPoint)) {
                            return null;
                        }
                        current = new List<OfficePoint> { currentPoint };
                        subpathStart = currentPoint;
                        hasCurrentPoint = true;
                        hasSubpathStart = true;
                        break;
                    case OfficePathCommandKind.LineTo:
                        if (!EnsureCurrentContour(
                                ref current,
                                currentPoint,
                                hasCurrentPoint,
                                ref subpathStart,
                                ref hasSubpathStart,
                                budget)) {
                            return null;
                        }
                        if (!TryTransformAndReserve(command.Point, transform, budget, out currentPoint)) {
                            return null;
                        }
                        current!.Add(currentPoint);
                        hasCurrentPoint = true;
                        break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        if (!EnsureCurrentContour(
                                ref current,
                                currentPoint,
                                hasCurrentPoint,
                                ref subpathStart,
                                ref hasSubpathStart,
                                budget) ||
                            !TryAddQuadraticPoints(
                                currentPoint,
                                command,
                                transform,
                                current!,
                                budget,
                                out currentPoint)) {
                            return null;
                        }
                        hasCurrentPoint = true;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        if (!EnsureCurrentContour(
                                ref current,
                                currentPoint,
                                hasCurrentPoint,
                                ref subpathStart,
                                ref hasSubpathStart,
                                budget) ||
                            !TryAddCubicPoints(
                                currentPoint,
                                command,
                                transform,
                                current!,
                                budget,
                                out currentPoint)) {
                            return null;
                        }
                        hasCurrentPoint = true;
                        break;
                    case OfficePathCommandKind.Close:
                        AddContour(contours, current, closed: true);
                        current = null;
                        if (hasSubpathStart) {
                            currentPoint = subpathStart;
                            hasCurrentPoint = true;
                        }
                        hasSubpathStart = false;
                        break;
                }
            }

            AddContour(contours, current, closed: false);
            return new VisualPath(contours, fillRule);
        }

        private static bool EnsureCurrentContour(
            ref List<OfficePoint>? current,
            OfficePoint currentPoint,
            bool hasCurrentPoint,
            ref OfficePoint subpathStart,
            ref bool hasSubpathStart,
            VisualGeometryBudget budget) {
            if (current != null) {
                return true;
            }
            if (!hasCurrentPoint || !budget.TryAddPoints(1)) {
                return false;
            }

            current = new List<OfficePoint> { currentPoint };
            subpathStart = currentPoint;
            hasSubpathStart = true;
            return true;
        }

        private static bool TryAddQuadraticPoints(
            OfficePoint transformedStart,
            OfficePathCommand command,
            OfficeTransform transform,
            List<OfficePoint> destination,
            VisualGeometryBudget budget,
            out OfficePoint transformedEnd) {
            transformedEnd = default;
            OfficePoint control = transform.TransformPoint(command.ControlPoint1);
            transformedEnd = transform.TransformPoint(command.Point);
            if (!IsFinite(control.X) ||
                !IsFinite(control.Y) ||
                !IsFinite(transformedEnd.X) ||
                !IsFinite(transformedEnd.Y) ||
                !budget.TryAddPoints(24)) {
                budget.Exhaust();
                return false;
            }

            destination.AddRange(OfficeGeometry.CreateQuadraticBezierPoints(
                transformedStart,
                control,
                transformedEnd,
                24));
            return true;
        }

        private static bool TryAddCubicPoints(
            OfficePoint transformedStart,
            OfficePathCommand command,
            OfficeTransform transform,
            List<OfficePoint> destination,
            VisualGeometryBudget budget,
            out OfficePoint transformedEnd) {
            transformedEnd = default;
            OfficePoint control1 = transform.TransformPoint(command.ControlPoint1);
            OfficePoint control2 = transform.TransformPoint(command.ControlPoint2);
            transformedEnd = transform.TransformPoint(command.Point);
            if (!IsFinite(control1.X) ||
                !IsFinite(control1.Y) ||
                !IsFinite(control2.X) ||
                !IsFinite(control2.Y) ||
                !IsFinite(transformedEnd.X) ||
                !IsFinite(transformedEnd.Y) ||
                !budget.TryAddPoints(24)) {
                budget.Exhaust();
                return false;
            }

            destination.AddRange(OfficeGeometry.CreateCubicBezierPoints(
                transformedStart,
                control1,
                control2,
                transformedEnd,
                24));
            return true;
        }

        private static bool TryTransformAndReserve(
            OfficePoint source,
            OfficeTransform transform,
            VisualGeometryBudget budget,
            out OfficePoint result) {
            result = transform.TransformPoint(source);
            if (!IsFinite(result.X) ||
                !IsFinite(result.Y) ||
                !budget.TryAddPoints(1)) {
                budget.Exhaust();
                return false;
            }

            return true;
        }

        private static void AddContour(
            List<VisualContour> contours,
            List<OfficePoint>? points,
            bool closed) {
            if (points == null || points.Count < 2) {
                return;
            }

            if (PointsEqual(points[0], points[points.Count - 1])) {
                points.RemoveAt(points.Count - 1);
                closed = true;
            }

            if (points.Count >= 2) {
                contours.Add(new VisualContour(points, closed));
            }
        }
    }

}
