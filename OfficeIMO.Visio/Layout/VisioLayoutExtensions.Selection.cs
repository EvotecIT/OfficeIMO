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

            foreach (VisioShape shape in selection) {
                switch (alignment) {
                    case VisioHorizontalAlignment.Left:
                        shape.PinX = bounds.Left + shape.Width / 2D;
                        break;
                    case VisioHorizontalAlignment.Center:
                        shape.PinX = bounds.CenterX;
                        break;
                    case VisioHorizontalAlignment.Right:
                        shape.PinX = bounds.Right - shape.Width / 2D;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(alignment));
                }
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

            foreach (VisioShape shape in selection) {
                switch (alignment) {
                    case VisioVerticalAlignment.Bottom:
                        shape.PinY = bounds.Bottom + shape.Height / 2D;
                        break;
                    case VisioVerticalAlignment.Middle:
                        shape.PinY = bounds.CenterY;
                        break;
                    case VisioVerticalAlignment.Top:
                        shape.PinY = bounds.Top - shape.Height / 2D;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(alignment));
                }
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

            if (selection.Count < 3) {
                return selection;
            }

            List<VisioShape> ordered;
            switch (axis) {
                case VisioDistributionAxis.Horizontal:
                    ordered = selection.OrderBy(shape => shape.PinX).ToList();
                    DistributeCenters(ordered, true);
                    break;
                case VisioDistributionAxis.Vertical:
                    ordered = selection.OrderBy(shape => shape.PinY).ToList();
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

            List<VisioShape> ordered = OrderSelection(selection, effectiveOptions.Order);
            int columns = ResolveColumnCount(effectiveOptions.Columns, ordered.Count);
            int rows = (int)Math.Ceiling(ordered.Count / (double)columns);
            double[] columnWidths = new double[columns];
            double[] rowHeights = new double[rows];

            for (int index = 0; index < ordered.Count; index++) {
                int row = index / columns;
                int column = index % columns;
                columnWidths[column] = Math.Max(columnWidths[column], ordered[index].Width);
                rowHeights[row] = Math.Max(rowHeights[row], ordered[index].Height);
            }

            VisioShapeBounds originalBounds = selection.GetShapeBounds();
            double startLeft = originalBounds.Left;
            double startTop = originalBounds.Top;
            if (!effectiveOptions.PreserveTopLeft) {
                VisioShape first = ordered[0];
                startLeft = first.PinX - first.Width / 2D;
                startTop = first.PinY + first.Height / 2D;
            }

            for (int index = 0; index < ordered.Count; index++) {
                int row = index / columns;
                int column = index % columns;
                VisioShape shape = ordered[index];
                double cellLeft = startLeft + SumBefore(columnWidths, column) + (effectiveOptions.HorizontalSpacing * column);
                double cellTop = startTop - SumBefore(rowHeights, row) - (effectiveOptions.VerticalSpacing * row);
                shape.PinX = cellLeft + columnWidths[column] / 2D;
                shape.PinY = cellTop - rowHeights[row] / 2D;
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

            if (options.HorizontalSpacing < 0) {
                throw new ArgumentOutOfRangeException(nameof(options.HorizontalSpacing), "Spacing cannot be negative.");
            }

            if (options.VerticalSpacing < 0) {
                throw new ArgumentOutOfRangeException(nameof(options.VerticalSpacing), "Spacing cannot be negative.");
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
                if (selectedShapes.Contains(connector.From) && selectedShapes.Contains(connector.To)) {
                    connector.RouteOrthogonal(style, (routeIndex % 3) * 0.04D);
                    routeIndex++;
                }
            }
        }

        private static void MoveShapes(IEnumerable<VisioShape> shapes, double deltaX, double deltaY) {
            foreach (VisioShape shape in shapes) {
                shape.PinX += deltaX;
                shape.PinY += deltaY;
            }
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
            VisioShape first = orderedShapes[0];
            VisioShape last = orderedShapes[orderedShapes.Count - 1];
            double firstCenter = horizontal ? first.PinX : first.PinY;
            double lastCenter = horizontal ? last.PinX : last.PinY;
            double step = (lastCenter - firstCenter) / (orderedShapes.Count - 1);

            for (int index = 1; index < orderedShapes.Count - 1; index++) {
                if (horizontal) {
                    orderedShapes[index].PinX = firstCenter + step * index;
                } else {
                    orderedShapes[index].PinY = firstCenter + step * index;
                }
            }
        }
    }
}
