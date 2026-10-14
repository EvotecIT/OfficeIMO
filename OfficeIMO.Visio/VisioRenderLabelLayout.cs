using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
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
        private readonly VisioRenderProjection _projection;
        private readonly IReadOnlyList<VisioShape> _shapes;
        private readonly IReadOnlyDictionary<VisioShape, VisioShapeBounds> _shapeBounds;
        private readonly IReadOnlyDictionary<VisioConnector, List<(double X, double Y)>> _connectorPaths;
        private readonly List<VisioShapeBounds> _placedLabels = new();

        private VisioRenderLabelLayout(VisioPage page, VisioRenderLayerVisibility layerVisibility, VisioRenderProjection projection) {
            _page = page;
            _projection = projection;
            var shapes = new List<VisioShape>();
            var bounds = new Dictionary<VisioShape, VisioShapeBounds>();
            foreach (VisioShape shape in page.AllShapes().Where(layerVisibility.IsVisible)) {
                try {
                    bounds.Add(shape, ToPhysicalBounds(GetPageShapeBounds(shape)));
                    shapes.Add(shape);
                } catch (Exception exception) when (exception is ArgumentException || exception is System.IO.InvalidDataException) {
                    // The renderer reports and omits invalid cached transforms. An
                    // omitted shape must not obstruct optional connector label layout.
                }
            }
            _shapes = shapes;
            _shapeBounds = bounds;
            _connectorPaths = page.Connectors.Where(connector => layerVisibility.IsVisible(connector) && VisioConnectorGeometry.HasVisibleLine(connector))
                .ToDictionary(connector => connector, connector => ToPhysicalPath(GetConnectorPoints(connector)));
        }

        internal static VisioRenderLabelLayout Create(VisioPage page, VisioRenderLayerVisibility layerVisibility) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return new VisioRenderLabelLayout(page, layerVisibility, VisioRenderProjection.Create(page));
        }

        internal VisioRenderConnectorLabelPlacement Resolve(VisioConnector connector, IReadOnlyList<(double X, double Y)> path) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (path == null) {
                throw new ArgumentNullException(nameof(path));
            }

            // Collision distances, clearances and scores share physical inches; callers retain drawing units.
            path = ToPhysicalPath(path);
            LabelPlacementSeed seed = CreateSeed(connector, path);
            LabelCandidate best = new LabelCandidate(seed.X, seed.Y, 0D, 0D);
            VisioShapeBounds bestBounds = GetBounds(best.X, best.Y, seed.Width, seed.Height, seed.LocPinX, seed.LocPinY);
            LabelScore bestScore = Score(connector, bestBounds, 0D, 0D);
            // A loaded native file supplies a stored position. Preview must not implicitly
            // run collision cleanup again; restored path intent still resolves endpoint edits.
            VisioConnectorLabelPlacement? stored = connector.LabelPlacement;
            bool absolute = stored?.AbsolutePinX.HasValue == true && stored.AbsolutePinY.HasValue;

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
            double ratio = _projection.DrawingToPhysical;
            return new VisioRenderConnectorLabelPlacement(best.X / ratio, best.Y / ratio, seed.Width / ratio, seed.Height / ratio, adjusted);
        }

        /// <summary>Resolves native label anchors and physical fitting defaults without collision adjustment.</summary>
        internal static VisioRenderConnectorLabelPlacement ResolveUnadjusted(VisioConnector connector,
            IReadOnlyList<(double X, double Y)> path, VisioRenderProjection projection) {
            VisioConnectorLabelPlacement? placement = VisioConnectorLabelFrame.ResolvePlacement(connector);
            double x, y;
            if (placement?.AbsolutePinX.HasValue == true && placement.AbsolutePinY.HasValue) {
                x = placement.AbsolutePinX.Value; y = placement.AbsolutePinY.Value;
            } else {
                (x, y) = OfficeGeometry.InterpolatePolyline(path, VisioConnectorLabelPlacement.ClampPosition(placement?.Position ?? 0.5D));
                x += placement?.OffsetX ?? 0D; y += placement?.OffsetY ?? 0D;
            }
            (double width, double height) = GetLabelDimensions(connector, projection);
            return new VisioRenderConnectorLabelPlacement(x, y, width, height, adjusted: false);
        }

        private static (double Width, double Height) GetLabelDimensions(VisioConnector connector, VisioRenderProjection projection) {
            double ratio = projection.DrawingToPhysical;
            double? width = connector.TextStyle?.TextWidth ?? VisioConnectorLabelFrame.ResolvePlacement(connector)?.Width;
            double? height = connector.TextStyle?.TextHeight ?? VisioConnectorLabelFrame.ResolvePlacement(connector)?.Height;
            // Authored/native boxes are drawing units. Only renderer fallbacks and fitting floors are physical inches.
            return (Math.Max(0.6D, width.HasValue ? width.Value * ratio : 1.35D) / ratio,
                Math.Max(0.18D, height.HasValue ? height.Value * ratio : 0.34D) / ratio);
        }

        private List<(double X, double Y)> ToPhysicalPath(IReadOnlyList<(double X, double Y)> path) =>
            path.Select(point => (point.X * _projection.DrawingToPhysical, point.Y * _projection.DrawingToPhysical)).ToList();

        private VisioShapeBounds ToPhysicalBounds(VisioShapeBounds bounds) =>
            new(bounds.Left * _projection.DrawingToPhysical, bounds.Bottom * _projection.DrawingToPhysical,
                bounds.Right * _projection.DrawingToPhysical, bounds.Top * _projection.DrawingToPhysical);

        private LabelPlacementSeed CreateSeed(VisioConnector connector, IReadOnlyList<(double X, double Y)> path) {
            VisioConnectorLabelPlacement? placement = VisioConnectorLabelFrame.ResolvePlacement(connector);
            (double drawingWidth, double drawingHeight) = GetLabelDimensions(connector, _projection);
            double ratio = _projection.DrawingToPhysical;
            double width = drawingWidth * ratio, height = drawingHeight * ratio;
            (double centerOffsetX, double centerOffsetY) = VisioConnectorGeometry.GetLabelCenter(connector, 0D, 0D, drawingWidth, drawingHeight);
            centerOffsetX *= ratio; centerOffsetY *= ratio;
            double locPinX = width / 2D - centerOffsetX;
            double locPinY = height / 2D - centerOffsetY;

            if (placement?.AbsolutePinX.HasValue == true && placement.AbsolutePinY.HasValue) {
                return new LabelPlacementSeed(
                    placement.AbsolutePinX.Value * ratio,
                    placement.AbsolutePinY.Value * ratio,
                    VisioConnectorLabelPlacement.ClampPosition(placement.Position),
                    width,
                    height,
                    locPinX,
                    locPinY,
                    placement.OffsetX * ratio,
                    placement.OffsetY * ratio);
            }

            double position = VisioConnectorLabelPlacement.ClampPosition(placement?.Position ?? 0.5D);
            (double x, double y) = OfficeGeometry.InterpolatePolyline(path, position);
            double offsetX = (placement?.OffsetX ?? 0D) * ratio;
            double offsetY = (placement?.OffsetY ?? 0D) * ratio;
            return new LabelPlacementSeed(x + offsetX, y + offsetY, position, width, height, locPinX, locPinY, offsetX, offsetY);
        }

        private IEnumerable<LabelCandidate> EnumerateCandidates(LabelPlacementSeed seed, IReadOnlyList<(double X, double Y)> path) {
            yield return new LabelCandidate(seed.X, seed.Y, 0D, 0D);

            for (int shift = 1; shift <= MaxPositionShifts; shift++) {
                double delta = shift * PositionStep;
                foreach (int direction in new[] { 1, -1 }) {
                    double positionDelta = delta * direction;
                    double position = VisioConnectorLabelPlacement.ClampPosition(seed.Position + positionDelta);
                    (double x, double y) = OfficeGeometry.InterpolatePolyline(path, position);
                    double candidateX = x + seed.OffsetX;
                    double candidateY = y + seed.OffsetY;
                    yield return new LabelCandidate(
                        candidateX,
                        candidateY,
                        OfficeGeometry.Distance(candidateX, candidateY, seed.X, seed.Y),
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
                if (ReferenceEquals(otherConnector, connector)) {
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
            double right = Math.Max(0D, bounds.Right - _projection.WidthInches);
            double top = Math.Max(0D, bounds.Top - _projection.HeightInches);
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

            return OfficeGeometry.SegmentIntersectsRectangle(
                first,
                second,
                bounds.Left,
                bounds.Bottom,
                bounds.Right,
                bounds.Top);
        }

        private static List<(double X, double Y)> GetConnectorPoints(VisioConnector connector) {
            return VisioConnectorGeometry.GetPoints(connector);
        }


        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            OfficePoint point = VisioNativeShapeTransform.Create(shape).PagePoint(x, y);
            return (point.X, point.Y);
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

        private static VisioShapeBounds GetPageShapeBounds(VisioShape shape) {
            (double left, double bottom, double right, double top) = GetPageBounds(shape);
            return new VisioShapeBounds(left, bottom, right, top);
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
            OfficeGeometry.ResolveRectangleBoundaryEndpoint(
                sourceLeft,
                sourceBottom,
                sourceRight,
                sourceTop,
                targetLeft,
                targetBottom,
                targetRight,
                targetTop,
                out x,
                out y);
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
