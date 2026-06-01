using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    internal sealed class VisioRenderLabelLayout {
        private const double SearchStep = 0.18D;
        private const int MaxSearchRings = 10;
        private const double PositionStep = 0.08D;
        private const int MaxPositionShifts = 4;
        private const double EndpointShapeOverlapWeight = 0.65D;
        private const double ShapeClearance = 0.02D;
        private const double LabelClearance = 0.04D;
        private const double ConnectorLineClearance = 0.03D;

        private readonly VisioPage _page;
        private readonly IReadOnlyList<VisioShape> _shapes;
        private readonly IReadOnlyDictionary<VisioShape, VisioShapeBounds> _shapeBounds;
        private readonly IReadOnlyDictionary<VisioConnector, List<(double X, double Y)>> _connectorPaths;
        private readonly List<VisioShapeBounds> _placedLabels = new();

        private VisioRenderLabelLayout(VisioPage page) {
            _page = page;
            _shapes = page.AllShapes();
            _shapeBounds = _shapes.ToDictionary(shape => shape, shape => shape.GetShapeBounds());
            _connectorPaths = page.Connectors.ToDictionary(connector => connector, GetConnectorPoints);
        }

        internal static VisioRenderLabelLayout Create(VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return new VisioRenderLabelLayout(page);
        }

        internal VisioRenderConnectorLabelPlacement Resolve(VisioConnector connector, IReadOnlyList<(double X, double Y)> path) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (path == null) {
                throw new ArgumentNullException(nameof(path));
            }

            LabelPlacementSeed seed = CreateSeed(connector, path);
            LabelCandidate best = new LabelCandidate(seed.X, seed.Y, 0D, 0D);
            VisioShapeBounds bestBounds = GetBounds(best.X, best.Y, seed.Width, seed.Height, seed.LocPinX, seed.LocPinY);
            LabelScore bestScore = Score(connector, bestBounds, 0D, 0D);
            bool absolute = connector.LabelPlacement?.AbsolutePinX.HasValue == true &&
                            connector.LabelPlacement.AbsolutePinY.HasValue;

            if (!absolute) {
                foreach (LabelCandidate candidate in EnumerateCandidates(seed, path)) {
                    VisioShapeBounds bounds = GetBounds(candidate.X, candidate.Y, seed.Width, seed.Height, seed.LocPinX, seed.LocPinY);
                    LabelScore score = Score(connector, bounds, candidate.DistanceFromSeed, Math.Abs(candidate.PositionDelta));
                    if (score.IsBetterThan(bestScore)) {
                        best = candidate;
                        bestBounds = bounds;
                        bestScore = score;
                    }

                    if (!bestScore.HasVisibleCollision) {
                        break;
                    }
                }
            }

            _placedLabels.Add(bestBounds);
            bool adjusted = Math.Abs(best.X - seed.X) > 1e-9 || Math.Abs(best.Y - seed.Y) > 1e-9;
            return new VisioRenderConnectorLabelPlacement(best.X, best.Y, seed.Width, seed.Height, adjusted);
        }

        private LabelPlacementSeed CreateSeed(VisioConnector connector, IReadOnlyList<(double X, double Y)> path) {
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            double width = Math.Max(0.6D, connector.TextStyle?.TextWidth ?? placement?.Width ?? 1.35D);
            double height = Math.Max(0.18D, connector.TextStyle?.TextHeight ?? placement?.Height ?? 0.34D);
            double locPinX = placement?.GetLocPinX() ?? width / 2D;
            double locPinY = placement?.GetLocPinY() ?? height / 2D;

            if (placement?.AbsolutePinX.HasValue == true && placement.AbsolutePinY.HasValue) {
                return new LabelPlacementSeed(
                    placement.AbsolutePinX.Value,
                    placement.AbsolutePinY.Value,
                    VisioConnectorLabelPlacement.ClampPosition(placement.Position),
                    width,
                    height,
                    locPinX,
                    locPinY,
                    placement.OffsetX,
                    placement.OffsetY);
            }

            double position = VisioConnectorLabelPlacement.ClampPosition(placement?.Position ?? 0.5D);
            (double x, double y) = InterpolatePath(path, position);
            double offsetX = placement?.OffsetX ?? 0D;
            double offsetY = placement?.OffsetY ?? 0D;
            return new LabelPlacementSeed(x + offsetX, y + offsetY, position, width, height, locPinX, locPinY, offsetX, offsetY);
        }

        private IEnumerable<LabelCandidate> EnumerateCandidates(LabelPlacementSeed seed, IReadOnlyList<(double X, double Y)> path) {
            yield return new LabelCandidate(seed.X, seed.Y, 0D, 0D);

            for (int shift = 1; shift <= MaxPositionShifts; shift++) {
                double delta = shift * PositionStep;
                foreach (int direction in new[] { 1, -1 }) {
                    double positionDelta = delta * direction;
                    double position = VisioConnectorLabelPlacement.ClampPosition(seed.Position + positionDelta);
                    (double x, double y) = InterpolatePath(path, position);
                    double candidateX = x + seed.OffsetX;
                    double candidateY = y + seed.OffsetY;
                    yield return new LabelCandidate(
                        candidateX,
                        candidateY,
                        Distance(candidateX, candidateY, seed.X, seed.Y),
                        positionDelta);
                }
            }

            for (int ring = 1; ring <= MaxSearchRings; ring++) {
                double distance = ring * SearchStep;
                yield return new LabelCandidate(seed.X, seed.Y + distance, distance, 0D);
                yield return new LabelCandidate(seed.X, seed.Y - distance, distance, 0D);
                yield return new LabelCandidate(seed.X + distance, seed.Y, distance, 0D);
                yield return new LabelCandidate(seed.X - distance, seed.Y, distance, 0D);
                yield return new LabelCandidate(seed.X + distance, seed.Y + distance, distance * Math.Sqrt(2D), 0D);
                yield return new LabelCandidate(seed.X - distance, seed.Y + distance, distance * Math.Sqrt(2D), 0D);
                yield return new LabelCandidate(seed.X + distance, seed.Y - distance, distance * Math.Sqrt(2D), 0D);
                yield return new LabelCandidate(seed.X - distance, seed.Y - distance, distance * Math.Sqrt(2D), 0D);
            }
        }

        private LabelScore Score(VisioConnector connector, VisioShapeBounds bounds, double distanceFromSeed, double positionDelta) {
            double pageOverflow = OutsidePageAmount(bounds);
            double shapeOverlap = 0D;
            VisioShapeBounds shapeClearanceBounds = ExpandBounds(bounds, ShapeClearance);
            foreach (VisioShape shape in _shapes) {
                bool endpointShape = ReferenceEquals(shape, connector.From) || ReferenceEquals(shape, connector.To);

                if (!endpointShape &&
                    (shape.IsContainer || shape.IsBackgroundSurface || VisioSemanticUserCells.IsGeneratedDiagramAdornment(shape))) {
                    continue;
                }

                VisioShapeBounds shapeBounds = _shapeBounds[shape];
                if (!endpointShape && Contains(shapeBounds, bounds)) {
                    continue;
                }

                double overlap = OverlapArea(shapeClearanceBounds, shapeBounds);
                shapeOverlap += endpointShape ? overlap * EndpointShapeOverlapWeight : overlap;
            }

            double labelOverlap = 0D;
            VisioShapeBounds labelClearanceBounds = ExpandBounds(bounds, LabelClearance);
            foreach (VisioShapeBounds placed in _placedLabels) {
                labelOverlap += OverlapArea(labelClearanceBounds, placed);
            }

            double connectorOverlap = 0D;
            foreach (VisioConnector otherConnector in _page.Connectors) {
                if (ReferenceEquals(otherConnector, connector) || !HasVisibleConnectorLine(otherConnector)) {
                    continue;
                }

                if (!_connectorPaths.TryGetValue(otherConnector, out List<(double X, double Y)>? points) ||
                    points.Count < 2) {
                    continue;
                }

                VisioShapeBounds paddedBounds = ExpandBounds(bounds, Math.Max(otherConnector.LineWeight / 2D, 0.02D) + ConnectorLineClearance);
                for (int i = 1; i < points.Count; i++) {
                    if (SegmentIntersectsBounds(points[i - 1], points[i], paddedBounds)) {
                        connectorOverlap += Math.Max(otherConnector.LineWeight, 0.01D);
                    }
                }
            }

            return new LabelScore(pageOverflow, shapeOverlap, labelOverlap, connectorOverlap, distanceFromSeed, positionDelta);
        }

        private double OutsidePageAmount(VisioShapeBounds bounds) {
            if (bounds.IsEmpty) {
                return 0D;
            }

            double left = Math.Max(0D, -bounds.Left);
            double bottom = Math.Max(0D, -bounds.Bottom);
            double right = Math.Max(0D, bounds.Right - _page.Width);
            double top = Math.Max(0D, bounds.Top - _page.Height);
            return left + bottom + right + top;
        }

        private static VisioShapeBounds GetBounds(double x, double y, double width, double height, double locPinX, double locPinY) =>
            new VisioShapeBounds(x - locPinX, y - locPinY, x - locPinX + width, y - locPinY + height);

        private static VisioShapeBounds ExpandBounds(VisioShapeBounds bounds, double padding) =>
            new VisioShapeBounds(bounds.Left - padding, bounds.Bottom - padding, bounds.Right + padding, bounds.Top + padding);

        private static bool SegmentIntersectsBounds((double X, double Y) first, (double X, double Y) second, VisioShapeBounds bounds) {
            if (bounds.IsEmpty) {
                return false;
            }

            if (Math.Max(first.X, second.X) < bounds.Left ||
                Math.Min(first.X, second.X) > bounds.Right ||
                Math.Max(first.Y, second.Y) < bounds.Bottom ||
                Math.Min(first.Y, second.Y) > bounds.Top) {
                return false;
            }

            if (ContainsPoint(bounds, first) || ContainsPoint(bounds, second)) {
                return true;
            }

            (double X, double Y) bottomLeft = (bounds.Left, bounds.Bottom);
            (double X, double Y) bottomRight = (bounds.Right, bounds.Bottom);
            (double X, double Y) topRight = (bounds.Right, bounds.Top);
            (double X, double Y) topLeft = (bounds.Left, bounds.Top);
            return SegmentsIntersect(first, second, bottomLeft, bottomRight) ||
                   SegmentsIntersect(first, second, bottomRight, topRight) ||
                   SegmentsIntersect(first, second, topRight, topLeft) ||
                   SegmentsIntersect(first, second, topLeft, bottomLeft);
        }

        private static bool ContainsPoint(VisioShapeBounds bounds, (double X, double Y) point) =>
            point.X >= bounds.Left && point.X <= bounds.Right &&
            point.Y >= bounds.Bottom && point.Y <= bounds.Top;

        private static bool SegmentsIntersect(
            (double X, double Y) firstStart,
            (double X, double Y) firstEnd,
            (double X, double Y) secondStart,
            (double X, double Y) secondEnd) {
            double d1 = Direction(secondStart, secondEnd, firstStart);
            double d2 = Direction(secondStart, secondEnd, firstEnd);
            double d3 = Direction(firstStart, firstEnd, secondStart);
            double d4 = Direction(firstStart, firstEnd, secondEnd);

            if (((d1 > 0D && d2 < 0D) || (d1 < 0D && d2 > 0D)) &&
                ((d3 > 0D && d4 < 0D) || (d3 < 0D && d4 > 0D))) {
                return true;
            }

            return (Math.Abs(d1) <= 1e-9 && OnSegment(secondStart, secondEnd, firstStart)) ||
                   (Math.Abs(d2) <= 1e-9 && OnSegment(secondStart, secondEnd, firstEnd)) ||
                   (Math.Abs(d3) <= 1e-9 && OnSegment(firstStart, firstEnd, secondStart)) ||
                   (Math.Abs(d4) <= 1e-9 && OnSegment(firstStart, firstEnd, secondEnd));
        }

        private static double Direction((double X, double Y) start, (double X, double Y) end, (double X, double Y) point) =>
            ((point.X - start.X) * (end.Y - start.Y)) - ((point.Y - start.Y) * (end.X - start.X));

        private static bool OnSegment((double X, double Y) start, (double X, double Y) end, (double X, double Y) point) =>
            point.X >= Math.Min(start.X, end.X) - 1e-9 &&
            point.X <= Math.Max(start.X, end.X) + 1e-9 &&
            point.Y >= Math.Min(start.Y, end.Y) - 1e-9 &&
            point.Y <= Math.Max(start.Y, end.Y) + 1e-9;

        private static bool HasVisibleConnectorLine(VisioConnector connector) =>
            connector.LinePattern != 0 && connector.LineWeight > 0D && connector.LineColor.A > 0;

        private static List<(double X, double Y)> GetConnectorPoints(VisioConnector connector) {
            ComputeConnectorEndpoints(connector, out double startX, out double startY, out double endX, out double endY);
            List<(double X, double Y)> points = new() { (startX, startY) };
            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    points.Add((waypoint.X, waypoint.Y));
                }
            } else if (connector.Kind == ConnectorKind.RightAngle) {
                points.Add((startX, endY));
            }

            points.Add((endX, endY));
            return points;
        }

        private static void ComputeConnectorEndpoints(VisioConnector connector, out double startX, out double startY, out double endX, out double endY) {
            if (connector.FromConnectionPoint != null) {
                (startX, startY) = GetPagePoint(connector.From, connector.FromConnectionPoint.X, connector.FromConnectionPoint.Y);
            } else {
                (double fromLeft, double fromBottom, double fromRight, double fromTop) = GetPageBounds(connector.From);
                (double toLeft, double toBottom, double toRight, double toTop) = GetPageBounds(connector.To);
                ResolveFallbackEndpoint(fromLeft, fromBottom, fromRight, fromTop, toLeft, toBottom, toRight, toTop, out startX, out startY);
            }

            if (connector.ToConnectionPoint != null) {
                (endX, endY) = GetPagePoint(connector.To, connector.ToConnectionPoint.X, connector.ToConnectionPoint.Y);
            } else {
                (double toLeft, double toBottom, double toRight, double toTop) = GetPageBounds(connector.To);
                (double fromLeft, double fromBottom, double fromRight, double fromTop) = GetPageBounds(connector.From);
                ResolveFallbackEndpoint(toLeft, toBottom, toRight, toTop, fromLeft, fromBottom, fromRight, fromTop, out endX, out endY);
            }
        }

        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            (double absX, double absY) = shape.GetAbsolutePoint(x, y);
            return shape.Parent != null
                ? GetPagePoint(shape.Parent, absX, absY)
                : (absX, absY);
        }

        private static (double Left, double Bottom, double Right, double Top) GetPageBounds(VisioShape shape) {
            (double x1, double y1) = GetPagePoint(shape, 0, 0);
            (double x2, double y2) = GetPagePoint(shape, shape.Width, 0);
            (double x3, double y3) = GetPagePoint(shape, 0, shape.Height);
            (double x4, double y4) = GetPagePoint(shape, shape.Width, shape.Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return (left, bottom, right, top);
        }

        private static void ResolveFallbackEndpoint(
            double sourceLeft,
            double sourceBottom,
            double sourceRight,
            double sourceTop,
            double targetLeft,
            double targetBottom,
            double targetRight,
            double targetTop,
            out double x,
            out double y) {
            double sourceCenterX = (sourceLeft + sourceRight) / 2D;
            double sourceCenterY = (sourceBottom + sourceTop) / 2D;
            double targetCenterX = (targetLeft + targetRight) / 2D;
            double targetCenterY = (targetBottom + targetTop) / 2D;
            double dx = targetCenterX - sourceCenterX;
            double dy = targetCenterY - sourceCenterY;

            if (Math.Abs(dy) > Math.Abs(dx)) {
                x = sourceCenterX;
                y = dy >= 0D ? sourceTop : sourceBottom;
                return;
            }

            x = dx >= 0D ? sourceRight : sourceLeft;
            y = sourceCenterY;
        }

        private static (double X, double Y) InterpolatePath(IReadOnlyList<(double X, double Y)> points, double position) {
            if (points.Count == 0) {
                return (0D, 0D);
            }

            if (points.Count == 1) {
                return points[0];
            }

            double total = 0D;
            for (int i = 1; i < points.Count; i++) {
                total += Distance(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y);
            }

            if (total <= 0D) {
                return points[0];
            }

            double target = total * VisioConnectorLabelPlacement.ClampPosition(position);
            double traversed = 0D;
            for (int i = 1; i < points.Count; i++) {
                double segment = Distance(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y);
                if (segment <= 0D) {
                    continue;
                }

                if (traversed + segment >= target) {
                    double t = (target - traversed) / segment;
                    return (
                        points[i - 1].X + ((points[i].X - points[i - 1].X) * t),
                        points[i - 1].Y + ((points[i].Y - points[i - 1].Y) * t));
                }

                traversed += segment;
            }

            return points[points.Count - 1];
        }

        private static bool Contains(VisioShapeBounds outer, VisioShapeBounds inner) {
            const double tolerance = 1e-6;
            return outer.Left <= inner.Left + tolerance &&
                   outer.Bottom <= inner.Bottom + tolerance &&
                   outer.Right + tolerance >= inner.Right &&
                   outer.Top + tolerance >= inner.Top;
        }

        private static double OverlapArea(VisioShapeBounds first, VisioShapeBounds second) {
            if (first.IsEmpty || second.IsEmpty) {
                return 0D;
            }

            double width = Math.Max(0D, Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left));
            double height = Math.Max(0D, Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom));
            return width * height;
        }

        private static double Distance(double x1, double y1, double x2, double y2) {
            double dx = x2 - x1;
            double dy = y2 - y1;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }

        private readonly struct LabelPlacementSeed {
            public LabelPlacementSeed(double x, double y, double position, double width, double height, double locPinX, double locPinY, double offsetX, double offsetY) {
                X = x;
                Y = y;
                Position = position;
                Width = width;
                Height = height;
                LocPinX = locPinX;
                LocPinY = locPinY;
                OffsetX = offsetX;
                OffsetY = offsetY;
            }

            public double X { get; }

            public double Y { get; }

            public double Position { get; }

            public double Width { get; }

            public double Height { get; }

            public double LocPinX { get; }

            public double LocPinY { get; }

            public double OffsetX { get; }

            public double OffsetY { get; }
        }

        private readonly struct LabelCandidate {
            public LabelCandidate(double x, double y, double distanceFromSeed, double positionDelta) {
                X = x;
                Y = y;
                DistanceFromSeed = distanceFromSeed;
                PositionDelta = positionDelta;
            }

            public double X { get; }

            public double Y { get; }

            public double DistanceFromSeed { get; }

            public double PositionDelta { get; }
        }

        private readonly struct LabelScore {
            public LabelScore(double pageOverflow, double shapeOverlap, double labelOverlap, double connectorOverlap, double distanceFromSeed, double positionDelta) {
                PageOverflow = pageOverflow;
                ShapeOverlap = shapeOverlap;
                LabelOverlap = labelOverlap;
                ConnectorOverlap = connectorOverlap;
                DistanceFromSeed = distanceFromSeed;
                PositionDelta = positionDelta;
            }

            private double PageOverflow { get; }

            private double ShapeOverlap { get; }

            private double LabelOverlap { get; }

            private double ConnectorOverlap { get; }

            private double DistanceFromSeed { get; }

            private double PositionDelta { get; }

            public bool HasVisibleCollision => PageOverflow > 1e-6 || ShapeOverlap > 1e-6 || LabelOverlap > 1e-6 || ConnectorOverlap > 1e-6;

            public bool IsBetterThan(LabelScore other) {
                int collision = Compare(CollisionPenalty, other.CollisionPenalty);
                if (collision != 0) {
                    return collision < 0;
                }

                int distance = Compare(DistanceFromSeed, other.DistanceFromSeed);
                if (distance != 0) {
                    return distance < 0;
                }

                return Compare(PositionDelta, other.PositionDelta) < 0;
            }

            private double CollisionPenalty => (PageOverflow * 200D) + (ShapeOverlap * 800D) + (LabelOverlap * 1000D) + (ConnectorOverlap * 1200D);

            private static int Compare(double first, double second) {
                if (Math.Abs(first - second) < 1e-9) {
                    return 0;
                }

                return first < second ? -1 : 1;
            }
        }
    }

    internal readonly struct VisioRenderConnectorLabelPlacement {
        public VisioRenderConnectorLabelPlacement(double x, double y, double width, double height, bool adjusted) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Adjusted = adjusted;
        }

        public double X { get; }

        public double Y { get; }

        public double Width { get; }

        public double Height { get; }

        public bool Adjusted { get; }
    }
}
