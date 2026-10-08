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
        /// Gets the bounds of a shape selection.
        /// </summary>
        /// <param name="selection">Selection to inspect.</param>
        public static VisioShapeBounds GetShapeBounds(this VisioShapeSelection selection) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            return ((IEnumerable<VisioShape>)selection).GetShapeBounds();
        }

        /// <summary>
        /// Aligns selected shapes horizontally inside the current selection bounds.
        /// </summary>
        /// <param name="selection">Selection to align.</param>
        /// <param name="alignment">Horizontal alignment.</param>
        public static VisioShapeSelection Align(this VisioShapeSelection selection, VisioHorizontalAlignment alignment) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            VisioShapeBounds bounds = selection.GetShapeBounds();
            if (bounds.IsEmpty) {
                return selection;
            }

            foreach (VisioShape shape in selection.Distinct().OrderBy(GetShapeDepth)) {
                VisioShapeBounds current = shape.GetShapeBounds();
                double delta = alignment switch {
                    VisioHorizontalAlignment.Left => bounds.Left - current.Left,
                    VisioHorizontalAlignment.Center => bounds.CenterX - current.CenterX,
                    VisioHorizontalAlignment.Right => bounds.Right - current.Right,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
                VisioNativeShapeTransform.MoveInPage(shape, delta, 0);
            }

            return selection;
        }

        /// <summary>
        /// Aligns selected shapes vertically inside the current selection bounds.
        /// </summary>
        /// <param name="selection">Selection to align.</param>
        /// <param name="alignment">Vertical alignment.</param>
        public static VisioShapeSelection Align(this VisioShapeSelection selection, VisioVerticalAlignment alignment) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            VisioShapeBounds bounds = selection.GetShapeBounds();
            if (bounds.IsEmpty) {
                return selection;
            }

            foreach (VisioShape shape in selection.Distinct().OrderBy(GetShapeDepth)) {
                VisioShapeBounds current = shape.GetShapeBounds();
                double delta = alignment switch {
                    VisioVerticalAlignment.Bottom => bounds.Bottom - current.Bottom,
                    VisioVerticalAlignment.Middle => bounds.CenterY - current.CenterY,
                    VisioVerticalAlignment.Top => bounds.Top - current.Top,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
                VisioNativeShapeTransform.MoveInPage(shape, 0, delta);
            }

            return selection;
        }

        /// <summary>
        /// Distributes selected shapes by center point along the requested axis.
        /// </summary>
        /// <param name="selection">Selection to distribute.</param>
        /// <param name="axis">Distribution axis.</param>
        public static VisioShapeSelection Distribute(this VisioShapeSelection selection, VisioDistributionAxis axis) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            List<VisioShape> unique = selection.Distinct().ToList();
            if (unique.Count < 3) {
                return selection;
            }

            List<VisioShape> ordered;
            switch (axis) {
                case VisioDistributionAxis.Horizontal:
                    ordered = unique.OrderBy(shape => shape.GetShapeBounds().CenterX).ToList();
                    DistributeCenters(ordered, true);
                    break;
                case VisioDistributionAxis.Vertical:
                    ordered = unique.OrderBy(shape => shape.GetShapeBounds().CenterY).ToList();
                    DistributeCenters(ordered, false);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(axis));
            }

            return selection;
        }

        /// <summary>
        /// Distributes selected shapes horizontally by center point.
        /// </summary>
        /// <param name="selection">Selection to distribute.</param>
        public static VisioShapeSelection DistributeHorizontally(this VisioShapeSelection selection) {
            return selection.Distribute(VisioDistributionAxis.Horizontal);
        }

        /// <summary>
        /// Distributes selected shapes vertically by center point.
        /// </summary>
        /// <param name="selection">Selection to distribute.</param>
        public static VisioShapeSelection DistributeVertically(this VisioShapeSelection selection) {
            return selection.Distribute(VisioDistributionAxis.Vertical);
        }

        /// <summary>
        /// Relays out selected shapes into a deterministic grid and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="selection">Selection to relayout.</param>
        /// <param name="columns">Number of columns. When zero, OfficeIMO uses a near-square grid.</param>
        /// <param name="horizontalSpacing">Horizontal spacing between columns in inches.</param>
        /// <param name="verticalSpacing">Vertical spacing between rows in inches.</param>
        /// <param name="routeInternalConnectors">Whether connectors whose endpoints are both selected should be rerouted orthogonally.</param>
        public static VisioShapeSelection RelayoutAsGrid(this VisioShapeSelection selection, int columns = 0, double horizontalSpacing = 0.5D, double verticalSpacing = 0.5D, bool routeInternalConnectors = true) {
            return selection.RelayoutAsGrid(new VisioSelectionLayoutOptions {
                Columns = columns <= 0 ? null : columns,
                HorizontalSpacing = horizontalSpacing,
                VerticalSpacing = verticalSpacing,
                RouteInternalConnectors = routeInternalConnectors
            });
        }

        /// <summary>
        /// Relays out selected shapes into a deterministic grid and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="selection">Selection to relayout.</param>
        /// <param name="options">Layout options.</param>
        public static VisioShapeSelection RelayoutAsGrid(this VisioShapeSelection selection, VisioSelectionLayoutOptions? options) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            VisioSelectionLayoutOptions effectiveOptions = options ?? new VisioSelectionLayoutOptions();
            ValidateSelectionLayoutOptions(effectiveOptions);

            if (selection.Count == 0) {
                return selection;
            }

            List<VisioShape> ordered = OrderSelection(selection, effectiveOptions.Order).Distinct().ToList();
            int columns = ResolveColumnCount(effectiveOptions.Columns, ordered.Count);
            int rows = (int)Math.Ceiling(ordered.Count / (double)columns);
            var footprints = ordered.ToDictionary(shape => shape, shape => shape.GetShapeBounds());
            double[] columnWidths = new double[columns];
            double[] rowHeights = new double[rows];

            for (int index = 0; index < ordered.Count; index++) {
                int row = index / columns;
                int column = index % columns;
                columnWidths[column] = Math.Max(columnWidths[column], footprints[ordered[index]].Width);
                rowHeights[row] = Math.Max(rowHeights[row], footprints[ordered[index]].Height);
            }

            VisioShapeBounds originalBounds = selection.GetShapeBounds();
            double startLeft = originalBounds.Left;
            double startTop = originalBounds.Top;
            if (!effectiveOptions.PreserveTopLeft) {
                VisioShape first = ordered[0];
                startLeft = footprints[first].Left;
                startTop = footprints[first].Top;
            }

            var placements = new Dictionary<VisioShape, (double X, double Y)>();
            for (int index = 0; index < ordered.Count; index++) {
                int row = index / columns;
                int column = index % columns;
                VisioShape shape = ordered[index];
                double cellLeft = startLeft + SumBefore(columnWidths, column) + (effectiveOptions.HorizontalSpacing * column);
                double cellTop = startTop - SumBefore(rowHeights, row) - (effectiveOptions.VerticalSpacing * row);
                placements[shape] = (cellLeft + columnWidths[column] / 2D, cellTop - rowHeights[row] / 2D);
            }
            foreach (VisioShape shape in ordered.OrderBy(GetShapeDepth)) {
                VisioShapeBounds current = shape.GetShapeBounds();
                VisioNativeShapeTransform.MoveInPage(shape, placements[shape].X - current.CenterX, placements[shape].Y - current.CenterY);
            }

            if (effectiveOptions.RouteInternalConnectors) {
                RerouteInternalConnectors(selection, effectiveOptions.ConnectorRouteStyle);
            }

            return selection;
        }

        /// <summary>
        /// Relays out selected shapes as a horizontal row and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="selection">Selection to relayout.</param>
        /// <param name="spacing">Horizontal spacing between shapes in inches.</param>
        /// <param name="routeInternalConnectors">Whether connectors whose endpoints are both selected should be rerouted orthogonally.</param>
        public static VisioShapeSelection RelayoutAsHorizontalStack(this VisioShapeSelection selection, double spacing = 0.5D, bool routeInternalConnectors = true) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            return selection.RelayoutAsGrid(selection.Count, spacing, 0D, routeInternalConnectors);
        }

        /// <summary>
        /// Relays out selected shapes as a vertical stack and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="selection">Selection to relayout.</param>
        /// <param name="spacing">Vertical spacing between shapes in inches.</param>
        /// <param name="routeInternalConnectors">Whether connectors whose endpoints are both selected should be rerouted orthogonally.</param>
        public static VisioShapeSelection RelayoutAsVerticalStack(this VisioShapeSelection selection, double spacing = 0.5D, bool routeInternalConnectors = true) {
            return selection.RelayoutAsGrid(1, 0D, spacing, routeInternalConnectors);
        }


        private static void ValidateSelectionLayoutOptions(VisioSelectionLayoutOptions options) {
            if (options.Columns.HasValue && options.Columns.Value < 0) {
                throw new ArgumentOutOfRangeException(nameof(options.Columns), "Column count cannot be negative.");
            }

            if (!IsFiniteNonNegativeSelectionLayoutValue(options.HorizontalSpacing)) {
                throw new ArgumentOutOfRangeException(nameof(options.HorizontalSpacing), "Spacing must be a finite non-negative number.");
            }

            if (!IsFiniteNonNegativeSelectionLayoutValue(options.VerticalSpacing)) {
                throw new ArgumentOutOfRangeException(nameof(options.VerticalSpacing), "Spacing must be a finite non-negative number.");
            }

            if (!Enum.IsDefined(typeof(VisioSelectionLayoutOrder), options.Order)) {
                throw new ArgumentOutOfRangeException(nameof(options.Order));
            }

            if (!Enum.IsDefined(typeof(VisioConnectorRouteStyle), options.ConnectorRouteStyle)) {
                throw new ArgumentOutOfRangeException(nameof(options.ConnectorRouteStyle));
            }
        }

        private static List<VisioShape> OrderSelection(VisioShapeSelection selection, VisioSelectionLayoutOrder order) {
            switch (order) {
                case VisioSelectionLayoutOrder.SelectionOrder:
                    return selection.ToList();
                case VisioSelectionLayoutOrder.TopLeftToBottomRight:
                    return selection
                        .OrderByDescending(shape => shape.GetShapeBounds().Top)
                        .ThenBy(shape => shape.GetShapeBounds().Left)
                        .ThenBy(shape => shape.Id, StringComparer.Ordinal)
                        .ToList();
                case VisioSelectionLayoutOrder.LeftTopToRightBottom:
                    return selection
                        .OrderBy(shape => shape.GetShapeBounds().Left)
                        .ThenByDescending(shape => shape.GetShapeBounds().Top)
                        .ThenBy(shape => shape.Id, StringComparer.Ordinal)
                        .ToList();
                default:
                    throw new ArgumentOutOfRangeException(nameof(order));
            }
        }

        private static int ResolveColumnCount(int? columns, int count) {
            if (count <= 0) {
                return 1;
            }

            if (columns.HasValue && columns.Value > 0) {
                return Math.Min(columns.Value, count);
            }

            return Math.Max(1, (int)Math.Ceiling(Math.Sqrt(count)));
        }

        private static bool IsFiniteNonNegativeSelectionLayoutValue(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value) && value >= 0D;
        }

        private static double SumBefore(IReadOnlyList<double> values, int exclusiveEnd) {
            double sum = 0D;
            for (int i = 0; i < exclusiveEnd; i++) {
                sum += values[i];
            }

            return sum;
        }

        private static void RerouteInternalConnectors(VisioShapeSelection selection, VisioConnectorRouteStyle style) {
            VisioPage? page = selection.OwnerPage;
            if (page == null) {
                return;
            }

            HashSet<VisioShape> selectedShapes = new(selection);
            int routeIndex = 0;
            foreach (VisioConnector connector in page.Connectors) {
                if (connector.From != null && connector.To != null &&
                    selectedShapes.Contains(connector.From) && selectedShapes.Contains(connector.To)) {
                    connector.RouteOrthogonal(style, (routeIndex % 3) * 0.04D);
                    routeIndex++;
                }
            }
        }

        private static int GetShapeDepth(VisioShape shape) {
            int depth = 0;
            for (VisioShape? parent = shape.Parent; parent != null; parent = parent.Parent) depth++;
            return depth;
        }

        private static void MoveShapes(IEnumerable<VisioShape> shapes, double deltaX, double deltaY) {
            foreach (VisioShape shape in shapes)
                VisioNativeShapeTransform.MoveInPage(shape, deltaX, deltaY);
        }

        private static void MoveConnectorPageCoordinates(IEnumerable<VisioConnector> connectors, double deltaX, double deltaY) {
            foreach (VisioConnector connector in connectors) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    waypoint.X += deltaX;
                    waypoint.Y += deltaY;
                }

                VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
                if (placement?.AbsolutePinX.HasValue == true) {
                    placement.AbsolutePinX += deltaX;
                }

                if (placement?.AbsolutePinY.HasValue == true) {
                    placement.AbsolutePinY += deltaY;
                }
            }
        }

        private static void DistributeCenters(IReadOnlyList<VisioShape> orderedShapes, bool horizontal) {
            var centers = orderedShapes.ToDictionary(shape => shape, shape => shape.GetShapeBounds());
            double firstCenter = horizontal ? centers[orderedShapes[0]].CenterX : centers[orderedShapes[0]].CenterY;
            double lastCenter = horizontal ? centers[orderedShapes[orderedShapes.Count - 1]].CenterX : centers[orderedShapes[orderedShapes.Count - 1]].CenterY;
            double step = (lastCenter - firstCenter) / (orderedShapes.Count - 1);
            var targets = orderedShapes.Select((shape, index) => new { Shape = shape, Center = firstCenter + step * index });
            foreach (var target in targets.OrderBy(item => GetShapeDepth(item.Shape))) {
                VisioShapeBounds current = target.Shape.GetShapeBounds();
                VisioNativeShapeTransform.MoveInPage(target.Shape,
                    horizontal ? target.Center - current.CenterX : 0,
                    horizontal ? 0 : target.Center - current.CenterY);
            }
        }
    }
}
