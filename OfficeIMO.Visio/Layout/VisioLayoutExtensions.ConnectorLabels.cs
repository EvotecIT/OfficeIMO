using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Layout and geometry helpers for Visio pages, shapes, and selections.
    /// </summary>
    public static partial class VisioLayoutExtensions {
        /// <summary>
        /// Moves connector label boxes away from page edges, unrelated shapes, and previously placed connector labels.
        /// </summary>
        /// <param name="page">Page whose connector labels should be adjusted.</param>
        /// <param name="step">Search step in page units, expressed in inches.</param>
        /// <param name="maxAttempts">Number of search rings to try around the current label position.</param>
        /// <param name="avoidShapes">Whether labels should avoid unrelated non-container shapes.</param>
        /// <param name="avoidLabels">Whether labels should avoid other connector labels.</param>
        /// <param name="preferEndpointZones">Whether labels should prefer common endpoint zones and avoid unrelated background zones.</param>
        /// <param name="avoidConnectorPaths">Whether labels should avoid unrelated connector paths.</param>
        /// <param name="positionStep">Connector path-position search step, from 0.0 to 1.0.</param>
        /// <param name="maxPositionShifts">Number of positive and negative connector path-position shifts to try.</param>
        /// <param name="optimizationPasses">Number of whole-page label optimization passes to run after the initial placement sweep.</param>
        public static VisioPage ResolveConnectorLabelOverlaps(this VisioPage page, double step = 0.18D, int maxAttempts = 12, bool avoidShapes = true, bool avoidLabels = true, bool preferEndpointZones = false, bool avoidConnectorPaths = true, double positionStep = 0.08D, int maxPositionShifts = 4, int optimizationPasses = 1) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (step <= 0D || double.IsNaN(step) || double.IsInfinity(step)) {
                throw new ArgumentOutOfRangeException(nameof(step), "Step must be a positive finite value.");
            }

            if (maxAttempts < 0) {
                throw new ArgumentOutOfRangeException(nameof(maxAttempts), "Attempt count cannot be negative.");
            }

            if (positionStep <= 0D || positionStep > 1D || double.IsNaN(positionStep) || double.IsInfinity(positionStep)) {
                throw new ArgumentOutOfRangeException(nameof(positionStep), "Position step must be a positive finite value no greater than 1.");
            }

            if (maxPositionShifts < 0) {
                throw new ArgumentOutOfRangeException(nameof(maxPositionShifts), "Position shift count cannot be negative.");
            }

            if (optimizationPasses < 1) {
                throw new ArgumentOutOfRangeException(nameof(optimizationPasses), "Optimization pass count must be at least one.");
            }

            IReadOnlyList<VisioShape> shapes = page.Shapes.ToList();
            Dictionary<VisioShape, VisioShapeBounds> shapeBounds = shapes.ToDictionary(shape => shape, shape => shape.GetShapeBounds());
            Dictionary<VisioConnector, List<Point>> connectorPaths = BuildConnectorPaths(page);
            List<ConnectorLabelBounds> placedLabels = new();

            foreach (VisioConnector connector in page.Connectors) {
                VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
                if (placement == null || string.IsNullOrWhiteSpace(connector.Label)) {
                    continue;
                }

                List<Point> path = connectorPaths[connector];
                if (!TryGetConnectorLabelBounds(connector, path, placement, out VisioShapeBounds currentBounds)) {
                    continue;
                }

                CandidateScore currentScore = ScoreConnectorLabel(page, connector, currentBounds, shapes, shapeBounds, placedLabels, connectorPaths, avoidShapes, avoidLabels, preferEndpointZones, avoidConnectorPaths);
                if (!currentScore.HasImprovementOpportunity) {
                    placedLabels.Add(new ConnectorLabelBounds(connector, currentBounds));
                    continue;
                }

                VisioConnectorLabelPlacement bestPlacement = placement.Clone();
                VisioShapeBounds bestBounds = currentBounds;
                CandidateScore bestScore = currentScore;

                foreach (LabelCandidate candidate in EnumerateLabelCandidates(maxAttempts, step, maxPositionShifts, positionStep)) {
                    VisioConnectorLabelPlacement candidatePlacement = CreateCandidatePlacement(placement, candidate);
                    if (!TryGetConnectorLabelBounds(connector, path, candidatePlacement, out VisioShapeBounds candidateBounds)) {
                        continue;
                    }

                    CandidateScore candidateScore = ScoreConnectorLabel(page, connector, candidateBounds, shapes, shapeBounds, placedLabels, connectorPaths, avoidShapes, avoidLabels, preferEndpointZones, avoidConnectorPaths);
                    if (candidateScore.IsBetterThan(bestScore)) {
                        bestPlacement = candidatePlacement;
                        bestBounds = candidateBounds;
                        bestScore = candidateScore;
                    }

                    if (!candidateScore.HasImprovementOpportunity) {
                        break;
                    }
                }

                connector.LabelPlacement = bestPlacement;
                placedLabels.Add(new ConnectorLabelBounds(connector, bestBounds));
            }

            if (avoidLabels && page.Connectors.Count > 1) {
                ResolveConnectorLabelGlobalOverlaps(page, step, maxAttempts, positionStep, maxPositionShifts, optimizationPasses, shapes, shapeBounds, connectorPaths, avoidShapes, preferEndpointZones, avoidConnectorPaths);
            }

            return page;
        }

        private static void ResolveConnectorLabelGlobalOverlaps(
            VisioPage page,
            double step,
            int maxAttempts,
            double positionStep,
            int maxPositionShifts,
            int optimizationPasses,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> shapeBounds,
            IReadOnlyDictionary<VisioConnector, List<Point>> connectorPaths,
            bool avoidShapes,
            bool preferEndpointZones,
            bool avoidConnectorPaths) {
            for (int pass = 0; pass < optimizationPasses; pass++) {
                List<ConnectorLabelBounds> labelBounds = GetConnectorLabelBounds(page);
                if (labelBounds.Count < 2) {
                    return;
                }

                CandidateScore before = ScorePageConnectorLabels(page, labelBounds, shapes, shapeBounds, connectorPaths, avoidShapes, preferEndpointZones, avoidConnectorPaths);
                IReadOnlyList<VisioConnector> connectors = optimizationPasses > 1
                    ? OrderConnectorsForLabelCleanup(page, labelBounds, shapes, shapeBounds, connectorPaths, avoidShapes, preferEndpointZones, avoidConnectorPaths)
                    : page.Connectors.ToList();
                foreach (VisioConnector connector in connectors) {
                    VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
                    if (placement == null || string.IsNullOrWhiteSpace(connector.Label)) {
                        continue;
                    }

                    List<Point> path = connectorPaths[connector];
                    if (!TryGetConnectorLabelBounds(connector, path, placement, out VisioShapeBounds currentBounds)) {
                        continue;
                    }

                    List<ConnectorLabelBounds> otherLabels = labelBounds
                        .Where(label => !ReferenceEquals(label.Connector, connector))
                        .ToList();
                    CandidateScore currentScore = ScoreConnectorLabel(page, connector, currentBounds, shapes, shapeBounds, otherLabels, connectorPaths, avoidShapes, avoidLabels: true, preferEndpointZones: preferEndpointZones, avoidConnectorPaths: avoidConnectorPaths);
                    if (!currentScore.HasImprovementOpportunity) {
                        continue;
                    }

                    VisioConnectorLabelPlacement bestPlacement = placement.Clone();
                    VisioShapeBounds bestBounds = currentBounds;
                    CandidateScore bestScore = currentScore;
                    foreach (LabelCandidate candidate in EnumerateLabelCandidates(maxAttempts, step, maxPositionShifts, positionStep)) {
                        VisioConnectorLabelPlacement candidatePlacement = CreateCandidatePlacement(placement, candidate);
                        if (!TryGetConnectorLabelBounds(connector, path, candidatePlacement, out VisioShapeBounds candidateBounds)) {
                            continue;
                        }

                        CandidateScore candidateScore = ScoreConnectorLabel(page, connector, candidateBounds, shapes, shapeBounds, otherLabels, connectorPaths, avoidShapes, avoidLabels: true, preferEndpointZones: preferEndpointZones, avoidConnectorPaths: avoidConnectorPaths);
                        if (candidateScore.IsBetterThan(bestScore)) {
                            bestPlacement = candidatePlacement;
                            bestBounds = candidateBounds;
                            bestScore = candidateScore;
                        }

                        if (!candidateScore.HasImprovementOpportunity) {
                            break;
                        }
                    }

                    connector.LabelPlacement = bestPlacement;
                    labelBounds = labelBounds
                        .Where(label => !ReferenceEquals(label.Connector, connector))
                        .Concat(new[] { new ConnectorLabelBounds(connector, bestBounds) })
                        .ToList();
                }

                CandidateScore after = ScorePageConnectorLabels(page, GetConnectorLabelBounds(page), shapes, shapeBounds, connectorPaths, avoidShapes, preferEndpointZones, avoidConnectorPaths);
                if (!after.HasImprovementOpportunity || !after.IsBetterThan(before)) {
                    break;
                }
            }
        }

        private static IReadOnlyList<VisioConnector> OrderConnectorsForLabelCleanup(
            VisioPage page,
            IReadOnlyList<ConnectorLabelBounds> labelBounds,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> shapeBounds,
            IReadOnlyDictionary<VisioConnector, List<Point>> connectorPaths,
            bool avoidShapes,
            bool preferEndpointZones,
            bool avoidConnectorPaths) {
            List<ConnectorLabelWorkItem> workItems = new();
            int index = 0;
            foreach (VisioConnector connector in page.Connectors) {
                VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
                if (placement == null || string.IsNullOrWhiteSpace(connector.Label)) {
                    continue;
                }

                if (!connectorPaths.TryGetValue(connector, out List<Point>? path) ||
                    !TryGetConnectorLabelBounds(connector, path, placement, out VisioShapeBounds bounds)) {
                    continue;
                }

                List<ConnectorLabelBounds> otherLabels = labelBounds
                    .Where(label => !ReferenceEquals(label.Connector, connector))
                    .ToList();
                CandidateScore score = ScoreConnectorLabel(page, connector, bounds, shapes, shapeBounds, otherLabels, connectorPaths, avoidShapes, avoidLabels: true, preferEndpointZones: preferEndpointZones, avoidConnectorPaths: avoidConnectorPaths);
                workItems.Add(new ConnectorLabelWorkItem(connector, score, index++));
            }

            return workItems
                .OrderByDescending(item => item.Score.TotalPenalty)
                .ThenBy(item => item.Index)
                .Select(item => item.Connector)
                .ToList();
        }

        private static CandidateScore ScorePageConnectorLabels(
            VisioPage page,
            IReadOnlyList<ConnectorLabelBounds> labelBounds,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> shapeBounds,
            IReadOnlyDictionary<VisioConnector, List<Point>> connectorPaths,
            bool avoidShapes,
            bool preferEndpointZones,
            bool avoidConnectorPaths) {
            double pagePenalty = 0D;
            double shapeOverlap = 0D;
            double labelOverlap = 0D;
            double connectorPathOverlap = 0D;
            double zonePenalty = 0D;

            foreach (ConnectorLabelBounds label in labelBounds) {
                List<ConnectorLabelBounds> otherLabels = labelBounds
                    .Where(other => !ReferenceEquals(other.Connector, label.Connector))
                    .ToList();
                CandidateScore score = ScoreConnectorLabel(page, label.Connector, label.Bounds, shapes, shapeBounds, otherLabels, connectorPaths, avoidShapes, avoidLabels: true, preferEndpointZones: preferEndpointZones, avoidConnectorPaths: avoidConnectorPaths);
                pagePenalty += score.PagePenalty;
                shapeOverlap += score.ShapeOverlap;
                labelOverlap += score.LabelOverlap;
                connectorPathOverlap += score.ConnectorPathOverlap;
                zonePenalty += score.ZonePenalty;
            }

            return new CandidateScore(pagePenalty, shapeOverlap, labelOverlap, connectorPathOverlap, zonePenalty);
        }

        private static List<ConnectorLabelBounds> GetConnectorLabelBounds(VisioPage page) {
            List<ConnectorLabelBounds> labels = new();
            foreach (VisioConnector connector in page.Connectors) {
                if (string.IsNullOrWhiteSpace(connector.Label)) {
                    continue;
                }

                List<Point> path = BuildConnectorPath(connector);
                if (TryGetConnectorLabelBounds(connector, path, out VisioShapeBounds bounds)) {
                    labels.Add(new ConnectorLabelBounds(connector, bounds));
                }
            }

            return labels;
        }

        private static Dictionary<VisioConnector, List<Point>> BuildConnectorPaths(VisioPage page) {
            Dictionary<VisioConnector, List<Point>> paths = new();
            foreach (VisioConnector connector in page.Connectors) {
                paths[connector] = BuildConnectorPath(connector);
            }

            return paths;
        }


        private static VisioShapeBounds GetConnectorContentBounds(VisioConnector connector) {
            List<Point> path = BuildConnectorPath(connector);
            VisioShapeBounds bounds = GetPointBounds(path);
            if (TryGetConnectorLabelBounds(connector, path, out VisioShapeBounds labelBounds)) {
                bounds = Combine(bounds, labelBounds);
            }

            return bounds;
        }

        private static List<Point> BuildConnectorPath(VisioConnector connector) {
            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
            List<Point> points = new() {
                new Point(startX, startY)
            };

            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    points.Add(new Point(waypoint.X, waypoint.Y));
                }
            } else if (connector.Kind == ConnectorKind.RightAngle) {
                points.Add(new Point(startX, endY));
            }

            points.Add(new Point(endX, endY));
            return points;
        }

        private static void ResolveEndpoint(VisioShape shape, VisioShape other, VisioConnectionPoint? connectionPoint, out double x, out double y) {
            if (connectionPoint != null) {
                (x, y) = shape.GetAbsolutePoint(connectionPoint.X, connectionPoint.Y);
                return;
            }

            VisioShapeBounds shapeBounds = shape.GetShapeBounds();
            VisioShapeBounds otherBounds = other.GetShapeBounds();
            double dx = otherBounds.CenterX - shapeBounds.CenterX;
            double dy = otherBounds.CenterY - shapeBounds.CenterY;

            if (Math.Abs(dx) >= Math.Abs(dy)) {
                x = dx >= 0 ? shapeBounds.Right : shapeBounds.Left;
                y = shapeBounds.CenterY;
            } else {
                x = shapeBounds.CenterX;
                y = dy >= 0 ? shapeBounds.Top : shapeBounds.Bottom;
            }
        }

        private static VisioShapeBounds GetPointBounds(IReadOnlyList<Point> points) {
            if (points.Count == 0) {
                return VisioShapeBounds.Empty;
            }

            double left = points[0].X;
            double bottom = points[0].Y;
            double right = points[0].X;
            double top = points[0].Y;
            for (int i = 1; i < points.Count; i++) {
                left = Math.Min(left, points[i].X);
                bottom = Math.Min(bottom, points[i].Y);
                right = Math.Max(right, points[i].X);
                top = Math.Max(top, points[i].Y);
            }

            return new VisioShapeBounds(left, bottom, right, top);
        }

        private static bool TryGetConnectorLabelBounds(VisioConnector connector, IReadOnlyList<Point> path, out VisioShapeBounds bounds) {
            bounds = VisioShapeBounds.Empty;
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            return placement != null && TryGetConnectorLabelBounds(connector, path, placement, out bounds);
        }

        private static bool TryGetConnectorLabelBounds(VisioConnector connector, IReadOnlyList<Point> path, VisioConnectorLabelPlacement placement, out VisioShapeBounds bounds) {
            bounds = VisioShapeBounds.Empty;
            if (placement == null || path.Count == 0) {
                return false;
            }

            Point pin = placement.AbsolutePinX.HasValue && placement.AbsolutePinY.HasValue
                ? new Point(placement.AbsolutePinX.Value, placement.AbsolutePinY.Value)
                : ResolvePathPoint(path, placement.Position).Offset(placement.OffsetX, placement.OffsetY);
            double locPinX = placement.GetLocPinX();
            double locPinY = placement.GetLocPinY();
            bounds = new VisioShapeBounds(
                pin.X - locPinX,
                pin.Y - locPinY,
                pin.X - locPinX + placement.Width,
                pin.Y - locPinY + placement.Height);
            return true;
        }

        private static IEnumerable<LabelCandidate> EnumerateLabelCandidates(int maxAttempts, double step, int maxPositionShifts, double positionStep) {
            yield return new LabelCandidate(0D, 0D, 0D);
            for (int shift = 1; shift <= maxPositionShifts; shift++) {
                double delta = shift * positionStep;
                yield return new LabelCandidate(0D, 0D, delta);
                yield return new LabelCandidate(0D, 0D, -delta);
            }

            for (int ring = 1; ring <= maxAttempts; ring++) {
                double distance = ring * step;
                yield return new LabelCandidate(0D, distance, 0D);
                yield return new LabelCandidate(0D, -distance, 0D);
                yield return new LabelCandidate(distance, 0D, 0D);
                yield return new LabelCandidate(-distance, 0D, 0D);
                yield return new LabelCandidate(distance, distance, 0D);
                yield return new LabelCandidate(-distance, distance, 0D);
                yield return new LabelCandidate(distance, -distance, 0D);
                yield return new LabelCandidate(-distance, -distance, 0D);
            }
        }

        private static IEnumerable<ShapeCandidate> EnumerateShapeCandidates(int maxAttempts, double step) {
            yield return new ShapeCandidate(0D, 0D);
            for (int ring = 1; ring <= maxAttempts; ring++) {
                double distance = ring * step;
                yield return new ShapeCandidate(distance, 0D);
                yield return new ShapeCandidate(0D, distance);
                yield return new ShapeCandidate(0D, -distance);
                yield return new ShapeCandidate(-distance, 0D);
                yield return new ShapeCandidate(distance, distance);
                yield return new ShapeCandidate(distance, -distance);
                yield return new ShapeCandidate(-distance, distance);
                yield return new ShapeCandidate(-distance, -distance);
            }
        }

        private static VisioConnectorLabelPlacement CreateCandidatePlacement(VisioConnectorLabelPlacement source, LabelCandidate candidate) {
            VisioConnectorLabelPlacement placement = source.Clone();
            if (placement.AbsolutePinX.HasValue && placement.AbsolutePinY.HasValue) {
                placement.AbsolutePinX += candidate.OffsetX;
                placement.AbsolutePinY += candidate.OffsetY;
            } else {
                placement.Position = VisioConnectorLabelPlacement.ClampPosition(placement.Position + candidate.PositionDelta);
                placement.OffsetX += candidate.OffsetX;
                placement.OffsetY += candidate.OffsetY;
            }

            return placement;
        }

        private static CandidateScore ScoreConnectorLabel(
            VisioPage page,
            VisioConnector connector,
            VisioShapeBounds labelBounds,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> shapeBounds,
            IReadOnlyList<ConnectorLabelBounds> placedLabels,
            IReadOnlyDictionary<VisioConnector, List<Point>> connectorPaths,
            bool avoidShapes,
            bool avoidLabels,
            bool preferEndpointZones,
            bool avoidConnectorPaths) {
            double pagePenalty = OutsidePageAmount(labelBounds, page);
            double shapeOverlap = 0D;
            if (avoidShapes) {
                foreach (VisioShape shape in shapes) {
                    if (ReferenceEquals(shape, connector.From) || ReferenceEquals(shape, connector.To)) {
                        continue;
                    }

                    if (shape.IsContainer || shape.IsBackgroundSurface || VisioSemanticUserCells.IsGeneratedDiagramAdornment(shape)) {
                        continue;
                    }

                    VisioShapeBounds bounds = shapeBounds[shape];
                    if (Contains(bounds, labelBounds)) {
                        continue;
                    }

                    shapeOverlap += OverlapArea(labelBounds, bounds);
                }
            }

            double zonePenalty = preferEndpointZones
                ? ScoreConnectorLabelZonePreference(connector, labelBounds, shapes, shapeBounds)
                : 0D;
            double labelOverlap = 0D;
            if (avoidLabels) {
                foreach (ConnectorLabelBounds placedLabel in placedLabels) {
                    labelOverlap += OverlapArea(labelBounds, placedLabel.Bounds);
                }
            }

            double connectorPathOverlap = avoidConnectorPaths
                ? ScoreConnectorPathOverlap(connector, labelBounds, connectorPaths)
                : 0D;

            return new CandidateScore(pagePenalty, shapeOverlap, labelOverlap, connectorPathOverlap, zonePenalty);
        }

        private static double ScoreConnectorPathOverlap(
            VisioConnector connector,
            VisioShapeBounds labelBounds,
            IReadOnlyDictionary<VisioConnector, List<Point>> connectorPaths) {
            double overlap = 0D;
            foreach (KeyValuePair<VisioConnector, List<Point>> entry in connectorPaths) {
                if (ReferenceEquals(entry.Key, connector)) {
                    continue;
                }

                IReadOnlyList<Point> path = entry.Value;
                for (int i = 1; i < path.Count; i++) {
                    if (SegmentIntersectsBounds(path[i - 1], path[i], labelBounds)) {
                        overlap++;
                    }
                }
            }

            return overlap;
        }

        private static double ScoreConnectorLabelZonePreference(
            VisioConnector connector,
            VisioShapeBounds labelBounds,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> shapeBounds) {
            VisioShapeBounds fromBounds = connector.From.GetShapeBounds();
            VisioShapeBounds toBounds = connector.To.GetShapeBounds();
            List<VisioShapeBounds> commonEndpointZones = new();
            double unrelatedZoneOverlap = 0D;

            foreach (VisioShape shape in shapes) {
                if (!shape.IsBackgroundSurface) {
                    continue;
                }

                VisioShapeBounds bounds = shapeBounds[shape];
                if (bounds.IsEmpty) {
                    continue;
                }

                bool containsFrom = Contains(bounds, fromBounds);
                bool containsTo = Contains(bounds, toBounds);
                if (containsFrom && containsTo) {
                    commonEndpointZones.Add(bounds);
                    continue;
                }

                if (!containsFrom && !containsTo) {
                    unrelatedZoneOverlap += OverlapArea(labelBounds, bounds);
                }
            }

            double commonZonePenalty = 0D;
            if (commonEndpointZones.Count > 0) {
                commonZonePenalty = commonEndpointZones
                    .Min(zone => OutsideContainerAmount(labelBounds, zone));
            }

            return unrelatedZoneOverlap + commonZonePenalty;
        }

        private static double OutsidePageAmount(VisioShapeBounds bounds, VisioPage page) {
            if (bounds.IsEmpty) {
                return 0D;
            }

            double left = Math.Max(0D, -bounds.Left);
            double bottom = Math.Max(0D, -bounds.Bottom);
            double right = Math.Max(0D, bounds.Right - page.Width);
            double top = Math.Max(0D, bounds.Top - page.Height);
            return left + bottom + right + top;
        }

        private static double OutsideContainerAmount(VisioShapeBounds inner, VisioShapeBounds outer) {
            if (inner.IsEmpty || outer.IsEmpty) {
                return 0D;
            }

            double left = Math.Max(0D, outer.Left - inner.Left);
            double bottom = Math.Max(0D, outer.Bottom - inner.Bottom);
            double right = Math.Max(0D, inner.Right - outer.Right);
            double top = Math.Max(0D, inner.Top - outer.Top);
            return left + bottom + right + top;
        }

        private static bool SegmentIntersectsBounds(Point a, Point b, VisioShapeBounds bounds) {
            if (bounds.IsEmpty) {
                return false;
            }

            if (PointInside(a, bounds) || PointInside(b, bounds)) {
                return true;
            }

            Point bottomLeft = new(bounds.Left, bounds.Bottom);
            Point bottomRight = new(bounds.Right, bounds.Bottom);
            Point topLeft = new(bounds.Left, bounds.Top);
            Point topRight = new(bounds.Right, bounds.Top);

            return SegmentsIntersect(a, b, bottomLeft, bottomRight) ||
                   SegmentsIntersect(a, b, bottomRight, topRight) ||
                   SegmentsIntersect(a, b, topRight, topLeft) ||
                   SegmentsIntersect(a, b, topLeft, bottomLeft);
        }

        private static bool PointInside(Point point, VisioShapeBounds bounds) {
            return point.X > bounds.Left && point.X < bounds.Right &&
                   point.Y > bounds.Bottom && point.Y < bounds.Top;
        }

        private static bool SegmentsIntersect(Point p1, Point p2, Point q1, Point q2) {
            double o1 = Orientation(p1, p2, q1);
            double o2 = Orientation(p1, p2, q2);
            double o3 = Orientation(q1, q2, p1);
            double o4 = Orientation(q1, q2, p2);

            if (o1 * o2 < 0D && o3 * o4 < 0D) {
                return true;
            }

            return IsZero(o1) && OnSegment(p1, q1, p2) ||
                   IsZero(o2) && OnSegment(p1, q2, p2) ||
                   IsZero(o3) && OnSegment(q1, p1, q2) ||
                   IsZero(o4) && OnSegment(q1, p2, q2);
        }

        private static double Orientation(Point a, Point b, Point c) {
            return ((b.X - a.X) * (c.Y - a.Y)) - ((b.Y - a.Y) * (c.X - a.X));
        }

        private static bool OnSegment(Point a, Point b, Point c) {
            return b.X >= Math.Min(a.X, c.X) - 1e-9 &&
                   b.X <= Math.Max(a.X, c.X) + 1e-9 &&
                   b.Y >= Math.Min(a.Y, c.Y) - 1e-9 &&
                   b.Y <= Math.Max(a.Y, c.Y) + 1e-9;
        }

        private static bool IsZero(double value) {
            return Math.Abs(value) < 1e-9;
        }

        private static Point ResolvePathPoint(IReadOnlyList<Point> points, double position) {
            double clampedPosition = VisioConnectorLabelPlacement.ClampPosition(position);
            double totalLength = 0D;
            for (int i = 1; i < points.Count; i++) {
                totalLength += Distance(points[i - 1], points[i]);
            }

            if (totalLength <= 0D) {
                return points[0];
            }

            double targetLength = totalLength * clampedPosition;
            double traversed = 0D;
            for (int i = 1; i < points.Count; i++) {
                Point from = points[i - 1];
                Point to = points[i];
                double segmentLength = Distance(from, to);
                if (segmentLength <= 0D) {
                    continue;
                }

                if (traversed + segmentLength >= targetLength) {
                    double segmentPosition = (targetLength - traversed) / segmentLength;
                    return new Point(
                        from.X + ((to.X - from.X) * segmentPosition),
                        from.Y + ((to.Y - from.Y) * segmentPosition));
                }

                traversed += segmentLength;
            }

            return points[points.Count - 1];
        }

        private static VisioShapeBounds Combine(VisioShapeBounds first, VisioShapeBounds second) {
            if (first.IsEmpty) {
                return second;
            }

            if (second.IsEmpty) {
                return first;
            }

            return new VisioShapeBounds(
                Math.Min(first.Left, second.Left),
                Math.Min(first.Bottom, second.Bottom),
                Math.Max(first.Right, second.Right),
                Math.Max(first.Top, second.Top));
        }

        private static bool Contains(VisioShapeBounds outer, VisioShapeBounds inner) {
            const double tolerance = 1e-6;
            return outer.Left <= inner.Left + tolerance &&
                   outer.Bottom <= inner.Bottom + tolerance &&
                   outer.Right + tolerance >= inner.Right &&
                   outer.Top + tolerance >= inner.Top;
        }

        private static double GetTotalShapeOverlap(VisioShape shape, IReadOnlyList<VisioShape> shapes) {
            VisioShapeBounds bounds = shape.GetShapeBounds();
            double total = 0D;
            foreach (VisioShape other in shapes) {
                if (ReferenceEquals(shape, other)) {
                    continue;
                }

                total += OverlapArea(bounds, other.GetShapeBounds());
            }

            return total;
        }

        private static double OverlapArea(VisioShapeBounds first, VisioShapeBounds second) {
            if (first.IsEmpty || second.IsEmpty) {
                return 0D;
            }

            double width = Math.Max(0D, Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left));
            double height = Math.Max(0D, Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom));
            return width * height;
        }

        private static double Distance(Point from, Point to) {
            double dx = to.X - from.X;
            double dy = to.Y - from.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }
    }
}
