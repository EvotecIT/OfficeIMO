using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private sealed partial class VisualPath {
        public static bool HasPositiveAreaIntersection(
            IReadOnlyList<VisualPath> paths,
            VisualGeometryBudget budget) {
            if (paths.Count == 0) {
                return false;
            }

            VisualBounds common = paths[0].Bounds;
            if (!common.HasPositiveArea) {
                return false;
            }

            for (int i = 1; i < paths.Count; i++) {
                if (!common.TryIntersectPositive(paths[i].Bounds, out common)) {
                    return false;
                }
            }

            if (TryCommonBoundsSamples(paths, common, budget)) {
                return true;
            }
            if (budget.Exceeded) {
                return true;
            }

            double sampleDistance = Math.Max(
                VisualGeometryEpsilon * 16D,
                Math.Min(common.Width, common.Height) * 0.000001D);
            for (int pathIndex = 0; pathIndex < paths.Count; pathIndex++) {
                VisualPath path = paths[pathIndex];
                for (int contourIndex = 0; contourIndex < path._contours.Count; contourIndex++) {
                    VisualContour contour = path._contours[contourIndex];
                    if (!contour.Bounds.IntersectsInclusive(common)) {
                        continue;
                    }

                    int segmentCount = contour.SegmentCount(closeForFill: true);
                    for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                        contour.GetSegment(segmentIndex, closeForFill: true, out OfficePoint start, out OfficePoint end);
                        if (TryBoundaryInteriorSamples(
                                start,
                                end,
                                sampleDistance,
                                paths,
                                budget)) {
                            return true;
                        }

                        if (budget.Exceeded) {
                            return true;
                        }
                    }
                }
            }

            for (int firstPathIndex = 0; firstPathIndex < paths.Count; firstPathIndex++) {
                VisualPath firstPath = paths[firstPathIndex];
                for (int secondPathIndex = firstPathIndex + 1; secondPathIndex < paths.Count; secondPathIndex++) {
                    VisualPath secondPath = paths[secondPathIndex];
                    if (!firstPath.Bounds.IntersectsInclusive(secondPath.Bounds)) {
                        continue;
                    }

                    for (int firstContourIndex = 0; firstContourIndex < firstPath._contours.Count; firstContourIndex++) {
                        VisualContour firstContour = firstPath._contours[firstContourIndex];
                        for (int secondContourIndex = 0; secondContourIndex < secondPath._contours.Count; secondContourIndex++) {
                            VisualContour secondContour = secondPath._contours[secondContourIndex];
                            if (!budget.TryUseOperation()) {
                                return true;
                            }
                            if (!firstContour.Bounds.IntersectsInclusive(secondContour.Bounds)) {
                                continue;
                            }

                            if (TryBoundaryIntersectionSamples(
                                    firstContour,
                                    secondContour,
                                    sampleDistance,
                                    paths,
                                    budget)) {
                                return true;
                            }

                            if (budget.Exceeded) {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private bool StrokeIntersectsSingleFill(
            VisualPath fill,
            double strokeHalfWidth,
            VisualGeometryBudget budget) {
            VisualBounds expandedBounds = Bounds.Expand(strokeHalfWidth);
            if (!expandedBounds.TryIntersectPositive(fill.Bounds, out _)) {
                return false;
            }

            double maximumDistanceSquared = strokeHalfWidth * strokeHalfWidth;
            if (!IsFinite(maximumDistanceSquared)) {
                return false;
            }

            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                if (!contour.Bounds.Expand(strokeHalfWidth).IntersectsInclusive(fill.Bounds)) {
                    continue;
                }

                int segmentCount = contour.SegmentCount(closeForFill: false);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    contour.GetSegment(segmentIndex, closeForFill: false, out OfficePoint start, out OfficePoint end);
                    if (TryStrokeInteriorSamples(
                            start,
                            end,
                            strokeHalfWidth,
                            new[] { fill },
                            budget)) {
                        return true;
                    }

                    VisualBounds segmentBounds = VisualBounds.FromSegment(start, end).Expand(strokeHalfWidth);
                    for (int fillContourIndex = 0; fillContourIndex < fill._contours.Count; fillContourIndex++) {
                        VisualContour fillContour = fill._contours[fillContourIndex];
                        if (!budget.TryUseOperation()) {
                            return true;
                        }
                        if (!segmentBounds.IntersectsInclusive(fillContour.Bounds)) {
                            continue;
                        }

                        int fillSegmentCount = fillContour.SegmentCount(closeForFill: true);
                        for (int fillSegmentIndex = 0; fillSegmentIndex < fillSegmentCount; fillSegmentIndex++) {
                            if (!budget.TryUseOperation()) {
                                return true;
                            }
                            fillContour.GetSegment(fillSegmentIndex, closeForFill: true, out OfficePoint fillStart, out OfficePoint fillEnd);
                            if (!VisualBounds.FromSegment(start, end)
                                    .Expand(strokeHalfWidth)
                                    .IntersectsInclusive(VisualBounds.FromSegment(fillStart, fillEnd))) {
                                continue;
                            }

                            if (SegmentDistanceSquared(start, end, fillStart, fillEnd) <
                                maximumDistanceSquared) {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private bool ContainsStrict(
            OfficePoint point,
            VisualGeometryBudget budget) {
            if (!Bounds.ContainsStrict(point)) {
                return false;
            }

            bool evenOddInside = false;
            int winding = 0;
            for (int contourIndex = 0; contourIndex < _contours.Count; contourIndex++) {
                VisualContour contour = _contours[contourIndex];
                if (!budget.TryUseOperation()) {
                    return false;
                }
                if (!contour.Bounds.ContainsInclusive(point)) {
                    continue;
                }

                int segmentCount = contour.SegmentCount(closeForFill: true);
                for (int segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
                    if (!budget.TryUseOperation()) {
                        return false;
                    }

                    contour.GetSegment(segmentIndex, closeForFill: true, out OfficePoint start, out OfficePoint end);
                    if (PointOnSegment(point, start, end)) {
                        return false;
                    }

                    if ((start.Y > point.Y) != (end.Y > point.Y) &&
                        point.X < ((end.X - start.X) *
                            (point.Y - start.Y) /
                            (end.Y - start.Y)) + start.X) {
                        evenOddInside = !evenOddInside;
                    }

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

            return FillRule == OfficeFillRule.EvenOdd
                ? evenOddInside
                : winding != 0;
        }

        private static bool TryCommonBoundsSamples(
            IReadOnlyList<VisualPath> paths,
            VisualBounds common,
            VisualGeometryBudget budget) {
            double centerX = (common.Left + common.Right) / 2D;
            double centerY = (common.Top + common.Bottom) / 2D;
            double quarterX = common.Width / 4D;
            double quarterY = common.Height / 4D;
            OfficePoint[] candidates = {
                new OfficePoint(centerX, centerY),
                new OfficePoint(centerX - quarterX, centerY - quarterY),
                new OfficePoint(centerX + quarterX, centerY - quarterY),
                new OfficePoint(centerX - quarterX, centerY + quarterY),
                new OfficePoint(centerX + quarterX, centerY + quarterY)
            };
            for (int i = 0; i < candidates.Length; i++) {
                if (AllContainStrict(paths, candidates[i], budget)) {
                    return true;
                }
                if (budget.Exceeded) {
                    return false;
                }
            }

            return false;
        }

        private static bool TryBoundaryInteriorSamples(
            OfficePoint start,
            OfficePoint end,
            double distance,
            IReadOnlyList<VisualPath> paths,
            VisualGeometryBudget budget) {
            double deltaX = end.X - start.X;
            double deltaY = end.Y - start.Y;
            double length = Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
            if (!IsFinite(length) || length <= VisualGeometryEpsilon) {
                return false;
            }

            double normalX = -deltaY / length * distance;
            double normalY = deltaX / length * distance;
            double[] fractions = { 0.25D, 0.5D, 0.75D };
            for (int i = 0; i < fractions.Length; i++) {
                double x = start.X + (deltaX * fractions[i]);
                double y = start.Y + (deltaY * fractions[i]);
                if (AllContainStrict(paths, new OfficePoint(x + normalX, y + normalY), budget) ||
                    AllContainStrict(paths, new OfficePoint(x - normalX, y - normalY), budget)) {
                    return true;
                }
                if (budget.Exceeded) {
                    return false;
                }
            }

            return false;
        }

        private static bool TryBoundaryIntersectionSamples(
            VisualContour first,
            VisualContour second,
            double distance,
            IReadOnlyList<VisualPath> paths,
            VisualGeometryBudget budget) {
            int firstSegmentCount = first.SegmentCount(closeForFill: true);
            int secondSegmentCount = second.SegmentCount(closeForFill: true);
            for (int firstSegmentIndex = 0; firstSegmentIndex < firstSegmentCount; firstSegmentIndex++) {
                first.GetSegment(firstSegmentIndex, closeForFill: true, out OfficePoint firstStart, out OfficePoint firstEnd);
                VisualBounds firstBounds = VisualBounds.FromSegment(firstStart, firstEnd);
                for (int secondSegmentIndex = 0; secondSegmentIndex < secondSegmentCount; secondSegmentIndex++) {
                    second.GetSegment(secondSegmentIndex, closeForFill: true, out OfficePoint secondStart, out OfficePoint secondEnd);
                    if (!budget.TryUseOperation()) {
                        return false;
                    }
                    if (!firstBounds.IntersectsInclusive(VisualBounds.FromSegment(secondStart, secondEnd))) {
                        continue;
                    }

                    if (TryGetSegmentIntersection(
                            firstStart,
                            firstEnd,
                            secondStart,
                            secondEnd,
                            out OfficePoint intersection) &&
                        TryIntersectionInteriorSamples(
                            firstStart,
                            firstEnd,
                            secondStart,
                            secondEnd,
                            intersection,
                            distance,
                            paths,
                            budget)) {
                        return true;
                    }

                    if (budget.Exceeded) {
                        return false;
                    }
                }
            }

            return false;
        }

        private static bool TryIntersectionInteriorSamples(
            OfficePoint firstStart,
            OfficePoint firstEnd,
            OfficePoint secondStart,
            OfficePoint secondEnd,
            OfficePoint intersection,
            double distance,
            IReadOnlyList<VisualPath> paths,
            VisualGeometryBudget budget) {
            if (!TryGetUnitNormal(firstStart, firstEnd, distance, out double firstX, out double firstY) ||
                !TryGetUnitNormal(secondStart, secondEnd, distance, out double secondX, out double secondY)) {
                return false;
            }

            OfficePoint[] candidates = {
                new OfficePoint(intersection.X + firstX + secondX, intersection.Y + firstY + secondY),
                new OfficePoint(intersection.X + firstX - secondX, intersection.Y + firstY - secondY),
                new OfficePoint(intersection.X - firstX + secondX, intersection.Y - firstY + secondY),
                new OfficePoint(intersection.X - firstX - secondX, intersection.Y - firstY - secondY),
                new OfficePoint(intersection.X + distance, intersection.Y),
                new OfficePoint(intersection.X - distance, intersection.Y),
                new OfficePoint(intersection.X, intersection.Y + distance),
                new OfficePoint(intersection.X, intersection.Y - distance)
            };
            for (int i = 0; i < candidates.Length; i++) {
                if (AllContainStrict(paths, candidates[i], budget)) {
                    return true;
                }
                if (budget.Exceeded) {
                    return false;
                }
            }

            return false;
        }

        private static bool TryStrokeInteriorSamples(
            OfficePoint start,
            OfficePoint end,
            double strokeHalfWidth,
            IReadOnlyList<VisualPath> fills,
            VisualGeometryBudget budget) {
            double deltaX = end.X - start.X;
            double deltaY = end.Y - start.Y;
            double length = Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
            if (!IsFinite(length) || length <= VisualGeometryEpsilon) {
                return false;
            }

            double normalDistance = strokeHalfWidth * 0.5D;
            double normalX = -deltaY / length * normalDistance;
            double normalY = deltaX / length * normalDistance;
            double[] fractions = { 0D, 0.25D, 0.5D, 0.75D, 1D };
            for (int i = 0; i < fractions.Length; i++) {
                double x = start.X + (deltaX * fractions[i]);
                double y = start.Y + (deltaY * fractions[i]);
                if (AllContainStrict(fills, new OfficePoint(x, y), budget) ||
                    AllContainStrict(fills, new OfficePoint(x + normalX, y + normalY), budget) ||
                    AllContainStrict(fills, new OfficePoint(x - normalX, y - normalY), budget)) {
                    return true;
                }
                if (budget.Exceeded) {
                    return false;
                }
            }

            return false;
        }

        private static bool AllContainStrict(
            IReadOnlyList<VisualPath> paths,
            OfficePoint point,
            VisualGeometryBudget budget) {
            for (int i = 0; i < paths.Count; i++) {
                if (!paths[i].ContainsStrict(point, budget)) {
                    return false;
                }
            }

            return true;
        }
    }
}
