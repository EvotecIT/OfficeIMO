using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Dependency-free visual quality checks for generated or loaded Visio diagrams.
    /// </summary>
    public static class VisioDiagramQualityAnalyzer {
        private const string ShapeOutsidePageKind = "ShapeOutsidePage";
        private const string ShapeOverlapKind = "ShapeOverlap";
        private const string ConnectorCrossesShapeKind = "ConnectorCrossesShape";
        private const string ConnectorLabelOutsidePageKind = "ConnectorLabelOutsidePage";
        private const string ConnectorLabelOverlapsShapeKind = "ConnectorLabelOverlapsShape";
        private const string ConnectorLabelOverlapKind = "ConnectorLabelOverlap";
        private const string ConnectorMissingLabelKind = "ConnectorMissingLabel";

        /// <summary>
        /// Analyzes every page in a document for common visual quality issues.
        /// </summary>
        /// <param name="document">Document to analyze.</param>
        /// <param name="options">Analysis options.</param>
        public static IReadOnlyList<VisioDiagramQualityIssue> AnalyzeVisualQuality(this VisioDocument document, VisioDiagramQualityOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            VisioDiagramQualityOptions resolvedOptions = options?.Clone() ?? new VisioDiagramQualityOptions();
            List<VisioDiagramQualityIssue> issues = new();
            foreach (VisioPage page in document.Pages) {
                issues.AddRange(page.AnalyzeVisualQuality(resolvedOptions));
            }

            return issues;
        }

        /// <summary>
        /// Creates a visual quality report for every page in a document.
        /// </summary>
        public static VisioDiagramQualityReport GetVisualQualityReport(this VisioDocument document, VisioDiagramQualityOptions? options = null) {
            return new VisioDiagramQualityReport(document.AnalyzeVisualQuality(options));
        }

        /// <summary>
        /// Throws when a document contains visual quality issues at or above the requested severity.
        /// </summary>
        public static VisioDocument EnsureVisualQuality(
            this VisioDocument document,
            VisioDiagramQualityOptions? options = null,
            VisioDiagramQualityIssueSeverity minimumSeverity = VisioDiagramQualityIssueSeverity.Warning) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            document.GetVisualQualityReport(options).EnsureClean(minimumSeverity);
            return document;
        }

        /// <summary>
        /// Analyzes a page for common visual quality issues.
        /// </summary>
        /// <param name="page">Page to analyze.</param>
        /// <param name="options">Analysis options.</param>
        public static IReadOnlyList<VisioDiagramQualityIssue> AnalyzeVisualQuality(this VisioPage page, VisioDiagramQualityOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioDiagramQualityOptions resolvedOptions = options?.Clone() ?? new VisioDiagramQualityOptions();
            List<VisioDiagramQualityIssue> issues = new();
            IReadOnlyList<VisioShape> shapes = GetAnalyzedShapes(page, resolvedOptions);
            Dictionary<VisioShape, VisioShapeBounds> boundsByShape = shapes.ToDictionary(shape => shape, shape => shape.GetShapeBounds());

            if (resolvedOptions.CheckPageBounds) {
                AnalyzePageBounds(page, boundsByShape, resolvedOptions, issues);
            }

            if (resolvedOptions.CheckShapeOverlaps) {
                AnalyzeShapeOverlaps(page, shapes, boundsByShape, resolvedOptions, issues);
            }

            if (resolvedOptions.CheckConnectorShapeIntersections || resolvedOptions.CheckConnectorLabels || resolvedOptions.RequireConnectorLabels) {
                AnalyzeConnectors(page, shapes, boundsByShape, resolvedOptions, issues);
            }

            return issues;
        }

        /// <summary>
        /// Creates a visual quality report for a page.
        /// </summary>
        public static VisioDiagramQualityReport GetVisualQualityReport(this VisioPage page, VisioDiagramQualityOptions? options = null) {
            return new VisioDiagramQualityReport(page.AnalyzeVisualQuality(options));
        }

        /// <summary>
        /// Throws when a page contains visual quality issues at or above the requested severity.
        /// </summary>
        public static VisioPage EnsureVisualQuality(
            this VisioPage page,
            VisioDiagramQualityOptions? options = null,
            VisioDiagramQualityIssueSeverity minimumSeverity = VisioDiagramQualityIssueSeverity.Warning) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            page.GetVisualQualityReport(options).EnsureClean(minimumSeverity);
            return page;
        }

        private static IReadOnlyList<VisioShape> GetAnalyzedShapes(VisioPage page, VisioDiagramQualityOptions options) {
            return options.IncludeGroupChildren
                ? page.AllShapes().ToList()
                : page.Shapes.ToList();
        }

        private static void AnalyzePageBounds(
            VisioPage page,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> boundsByShape,
            VisioDiagramQualityOptions options,
            List<VisioDiagramQualityIssue> issues) {
            double tolerance = Math.Max(0D, options.PageBoundsTolerance);
            foreach (KeyValuePair<VisioShape, VisioShapeBounds> entry in boundsByShape) {
                VisioShapeBounds bounds = entry.Value;
                if (bounds.Left < -tolerance ||
                    bounds.Bottom < -tolerance ||
                    bounds.Right > page.Width + tolerance ||
                    bounds.Top > page.Height + tolerance) {
                    issues.Add(new VisioDiagramQualityIssue(
                        VisioDiagramQualityIssueSeverity.Error,
                        ShapeOutsidePageKind,
                        $"Shape '{entry.Key.Id}' is outside the page bounds ({FormatBounds(bounds)} on {FormatSize(page.Width, page.Height)} page).",
                        page.Name,
                        entry.Key.Id));
                }
            }
        }

        private static void AnalyzeShapeOverlaps(
            VisioPage page,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> boundsByShape,
            VisioDiagramQualityOptions options,
            List<VisioDiagramQualityIssue> issues) {
            double minimumRatio = Math.Max(0D, options.MinimumShapeOverlapRatio);
            for (int i = 0; i < shapes.Count; i++) {
                VisioShape first = shapes[i];
                VisioShapeBounds firstBounds = boundsByShape[first];
                if (firstBounds.IsEmpty || firstBounds.Width <= 0 || firstBounds.Height <= 0) {
                    continue;
                }

                for (int j = i + 1; j < shapes.Count; j++) {
                    VisioShape second = shapes[j];
                    VisioShapeBounds secondBounds = boundsByShape[second];
                    if (secondBounds.IsEmpty || secondBounds.Width <= 0 || secondBounds.Height <= 0) {
                        continue;
                    }

                    if (IsBenignShapeOverlapPair(first, second)) {
                        continue;
                    }

                    if (options.IgnoreContainingShapeOverlaps &&
                        (Contains(firstBounds, secondBounds) || Contains(secondBounds, firstBounds))) {
                        continue;
                    }

                    double overlapArea = OverlapArea(firstBounds, secondBounds);
                    if (overlapArea <= 0D) {
                        continue;
                    }

                    double smallerArea = Math.Min(firstBounds.Width * firstBounds.Height, secondBounds.Width * secondBounds.Height);
                    double overlapRatio = smallerArea <= 0D ? 0D : overlapArea / smallerArea;
                    if (overlapRatio >= minimumRatio) {
                        issues.Add(new VisioDiagramQualityIssue(
                            VisioDiagramQualityIssueSeverity.Warning,
                            ShapeOverlapKind,
                            $"Shapes '{first.Id}' and '{second.Id}' overlap by {FormatPercent(overlapRatio)} of the smaller shape.",
                            page.Name,
                            first.Id,
                            second.Id));
                    }
                }
            }
        }

        private static void AnalyzeConnectors(
            VisioPage page,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> boundsByShape,
            VisioDiagramQualityOptions options,
            List<VisioDiagramQualityIssue> issues) {
            List<ConnectorLabelBounds> connectorLabelBounds = new();
            foreach (VisioConnector connector in page.Connectors) {
                List<Point> path = BuildConnectorPath(connector);
                if (options.RequireConnectorLabels && string.IsNullOrWhiteSpace(connector.Label)) {
                    issues.Add(new VisioDiagramQualityIssue(
                        VisioDiagramQualityIssueSeverity.Information,
                        ConnectorMissingLabelKind,
                        $"Connector '{connector.Id}' does not have a label.",
                        page.Name,
                        connectorId: connector.Id));
                }

                if (options.CheckConnectorLabels &&
                    !string.IsNullOrWhiteSpace(connector.Label) &&
                    TryGetConnectorLabelBounds(connector, path, out VisioShapeBounds labelBounds)) {
                    ConnectorLabelBounds connectorLabel = new(connector, labelBounds);
                    connectorLabelBounds.Add(connectorLabel);

                    if (IsOutsidePage(labelBounds, page, options.PageBoundsTolerance)) {
                        issues.Add(new VisioDiagramQualityIssue(
                            VisioDiagramQualityIssueSeverity.Warning,
                            ConnectorLabelOutsidePageKind,
                            $"Connector '{connector.Id}' label is outside the page bounds ({FormatBounds(labelBounds)}).",
                            page.Name,
                            connectorId: connector.Id));
                    }
                }

                if (!options.CheckConnectorShapeIntersections || !HasDeterministicRoute(connector)) {
                    continue;
                }

                foreach (VisioShape shape in shapes) {
                    if (ReferenceEquals(shape, connector.From) || ReferenceEquals(shape, connector.To)) {
                        continue;
                    }

                    if (IsConnectorIntersectionIgnoredShape(shape)) {
                        continue;
                    }

                    VisioShapeBounds shapeBounds = boundsByShape[shape];
                    if (options.IgnoreContainingShapeOverlaps &&
                        (Contains(shapeBounds, boundsByShape[connector.From]) || Contains(shapeBounds, boundsByShape[connector.To]))) {
                        continue;
                    }

                    if (PathIntersectsBounds(path, shapeBounds)) {
                        issues.Add(new VisioDiagramQualityIssue(
                            VisioDiagramQualityIssueSeverity.Warning,
                            ConnectorCrossesShapeKind,
                            $"Connector '{connector.Id}' crosses shape '{shape.Id}'.",
                            page.Name,
                            shape.Id,
                            connectorId: connector.Id));
                    }
                }
            }

            if (options.CheckConnectorLabels && options.CheckConnectorLabelShapeOverlaps) {
                AnalyzeConnectorLabelShapeOverlaps(page, shapes, boundsByShape, connectorLabelBounds, options, issues);
            }

            if (options.CheckConnectorLabels && options.CheckConnectorLabelOverlaps) {
                AnalyzeConnectorLabelOverlaps(page, connectorLabelBounds, options, issues);
            }
        }

        private static void AnalyzeConnectorLabelShapeOverlaps(
            VisioPage page,
            IReadOnlyList<VisioShape> shapes,
            IReadOnlyDictionary<VisioShape, VisioShapeBounds> boundsByShape,
            IReadOnlyList<ConnectorLabelBounds> connectorLabelBounds,
            VisioDiagramQualityOptions options,
            List<VisioDiagramQualityIssue> issues) {
            double minimumRatio = Math.Max(0D, options.MinimumConnectorLabelOverlapRatio);
            foreach (ConnectorLabelBounds label in connectorLabelBounds) {
                foreach (VisioShape shape in shapes) {
                    if (ReferenceEquals(shape, label.Connector.From) || ReferenceEquals(shape, label.Connector.To)) {
                        continue;
                    }

                    if (shape.IsBackgroundSurface || VisioSemanticUserCells.IsGeneratedDiagramAdornment(shape) || IsSequenceActivation(shape) || IsSequenceFragment(shape)) {
                        continue;
                    }

                    VisioShapeBounds shapeBounds = boundsByShape[shape];
                    if (shapeBounds.IsEmpty || label.Bounds.IsEmpty) {
                        continue;
                    }

                    if (options.IgnoreContainingShapeOverlaps && Contains(shapeBounds, label.Bounds)) {
                        continue;
                    }

                    double labelArea = BoundsArea(label.Bounds);
                    if (labelArea <= 0D) {
                        continue;
                    }

                    double overlapRatio = OverlapArea(label.Bounds, shapeBounds) / labelArea;
                    if (overlapRatio >= minimumRatio) {
                        issues.Add(new VisioDiagramQualityIssue(
                            VisioDiagramQualityIssueSeverity.Warning,
                            ConnectorLabelOverlapsShapeKind,
                            $"Connector '{label.Connector.Id}' label overlaps unrelated shape '{shape.Id}' by {FormatPercent(overlapRatio)} of the label.",
                            page.Name,
                            shape.Id,
                            connectorId: label.Connector.Id));
                    }
                }
            }
        }

        private static void AnalyzeConnectorLabelOverlaps(
            VisioPage page,
            IReadOnlyList<ConnectorLabelBounds> connectorLabelBounds,
            VisioDiagramQualityOptions options,
            List<VisioDiagramQualityIssue> issues) {
            double minimumRatio = Math.Max(0D, options.MinimumConnectorLabelOverlapRatio);
            for (int i = 0; i < connectorLabelBounds.Count; i++) {
                ConnectorLabelBounds first = connectorLabelBounds[i];
                for (int j = i + 1; j < connectorLabelBounds.Count; j++) {
                    ConnectorLabelBounds second = connectorLabelBounds[j];
                    double smallerArea = Math.Min(BoundsArea(first.Bounds), BoundsArea(second.Bounds));
                    if (smallerArea <= 0D) {
                        continue;
                    }

                    double overlapRatio = OverlapArea(first.Bounds, second.Bounds) / smallerArea;
                    if (overlapRatio >= minimumRatio) {
                        issues.Add(new VisioDiagramQualityIssue(
                            VisioDiagramQualityIssueSeverity.Warning,
                            ConnectorLabelOverlapKind,
                            $"Connector labels '{first.Connector.Id}' and '{second.Connector.Id}' overlap by {FormatPercent(overlapRatio)} of the smaller label.",
                            page.Name,
                            connectorId: first.Connector.Id,
                            otherConnectorId: second.Connector.Id));
                    }
                }
            }
        }

        private static bool IsBenignShapeOverlapPair(VisioShape first, VisioShape second) {
            return (first.IsBackgroundSurface && second.IsDiagramAdornment) ||
                   (second.IsBackgroundSurface && first.IsDiagramAdornment) ||
                   (first.IsBackgroundSurface && second.IsCallout) ||
                   (second.IsBackgroundSurface && first.IsCallout) ||
                   (IsSequenceActivation(first) && second.IsDiagramAdornment) ||
                   (IsSequenceActivation(second) && first.IsDiagramAdornment) ||
                   IsSequenceFragment(first) ||
                   IsSequenceFragment(second);
        }

        private static bool IsConnectorIntersectionIgnoredShape(VisioShape shape) {
            return shape.IsBackgroundSurface || VisioSemanticUserCells.IsGeneratedDiagramAdornment(shape) || IsSequenceActivation(shape) || IsSequenceFragment(shape);
        }

        private static bool IsSequenceActivation(VisioShape shape) {
            return string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.SequenceActivationKind, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSequenceFragment(VisioShape shape) {
            return string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.SequenceFragmentKind, StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasDeterministicRoute(VisioConnector connector) {
            return connector.Waypoints.Count > 0 ||
                   connector.Kind == ConnectorKind.RightAngle ||
                   connector.Kind == ConnectorKind.Straight ||
                   connector.Kind == ConnectorKind.Curved;
        }

        private static List<Point> BuildConnectorPath(VisioConnector connector) {
            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
            List<(double X, double Y)> waypoints = connector.Waypoints
                .Select(waypoint => (X: waypoint.X, Y: waypoint.Y))
                .ToList();

            return OfficeGeometry.BuildConnectorPolyline(
                    (startX, startY),
                    (endX, endY),
                    waypoints,
                    connector.Kind == ConnectorKind.RightAngle)
                .Select(point => new Point(point.X, point.Y))
                .ToList();
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

        private static bool TryGetConnectorLabelBounds(VisioConnector connector, IReadOnlyList<Point> path, out VisioShapeBounds bounds) {
            bounds = VisioShapeBounds.Empty;
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
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

        private static Point ResolvePathPoint(IReadOnlyList<Point> points, double position) {
            (double x, double y) = OfficeGeometry.InterpolatePolyline(
                points.Select(point => (X: point.X, Y: point.Y)).ToList(),
                position);
            return new Point(x, y);
        }

        private static bool PathIntersectsBounds(IReadOnlyList<Point> points, VisioShapeBounds bounds) {
            if (bounds.IsEmpty || points.Count < 2) {
                return false;
            }

            for (int i = 1; i < points.Count; i++) {
                if (SegmentIntersectsBounds(points[i - 1], points[i], bounds)) {
                    return true;
                }
            }

            return false;
        }

        private static bool SegmentIntersectsBounds(Point a, Point b, VisioShapeBounds bounds) {
            return OfficeGeometry.SegmentIntersectsRectangle(
                (a.X, a.Y),
                (b.X, b.Y),
                bounds.Left,
                bounds.Bottom,
                bounds.Right,
                bounds.Top);
        }

        private static bool IsOutsidePage(VisioShapeBounds bounds, VisioPage page, double tolerance) {
            double resolvedTolerance = Math.Max(0D, tolerance);
            return bounds.Left < -resolvedTolerance ||
                   bounds.Bottom < -resolvedTolerance ||
                   bounds.Right > page.Width + resolvedTolerance ||
                   bounds.Top > page.Height + resolvedTolerance;
        }

        private static bool Contains(VisioShapeBounds outer, VisioShapeBounds inner) {
            const double tolerance = 1e-6;
            return outer.Left <= inner.Left + tolerance &&
                   outer.Bottom <= inner.Bottom + tolerance &&
                   outer.Right + tolerance >= inner.Right &&
                   outer.Top + tolerance >= inner.Top;
        }

        private static double OverlapArea(VisioShapeBounds first, VisioShapeBounds second) {
            double width = Math.Max(0D, Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left));
            double height = Math.Max(0D, Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom));
            return width * height;
        }

        private static double BoundsArea(VisioShapeBounds bounds) {
            return bounds.IsEmpty ? 0D : Math.Max(0D, bounds.Width) * Math.Max(0D, bounds.Height);
        }

        private static bool IsZero(double value) {
            return Math.Abs(value) < 1e-9;
        }

        private static string FormatBounds(VisioShapeBounds bounds) {
            return string.Format(CultureInfo.InvariantCulture, "{0:0.###},{1:0.###},{2:0.###},{3:0.###}", bounds.Left, bounds.Bottom, bounds.Right, bounds.Top);
        }

        private static string FormatSize(double width, double height) {
            return string.Format(CultureInfo.InvariantCulture, "{0:0.###}x{1:0.###}", width, height);
        }

        private static string FormatPercent(double value) {
            return string.Format(CultureInfo.InvariantCulture, "{0:0.#}%", value * 100D);
        }

        private readonly struct Point {
            public Point(double x, double y) {
                X = x;
                Y = y;
            }

            public double X { get; }

            public double Y { get; }

            public Point Offset(double x, double y) {
                return new Point(X + x, Y + y);
            }
        }

        private readonly struct ConnectorLabelBounds {
            public ConnectorLabelBounds(VisioConnector connector, VisioShapeBounds bounds) {
                Connector = connector;
                Bounds = bounds;
            }

            public VisioConnector Connector { get; }

            public VisioShapeBounds Bounds { get; }
        }
    }
}
