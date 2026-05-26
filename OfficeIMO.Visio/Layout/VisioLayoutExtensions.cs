using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Layout and geometry helpers for Visio pages, shapes, and selections.
    /// </summary>
    public static class VisioLayoutExtensions {
        private const double MinimumPageSize = 0.1D;
        private const double DefaultHorizontalPadding = 0.25D;
        private const double DefaultVerticalPadding = 0.14D;

        /// <summary>
        /// Gets the page-space bounds for a shape.
        /// </summary>
        /// <param name="shape">Shape to inspect.</param>
        public static VisioShapeBounds GetShapeBounds(this VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            (double left, double bottom, double right, double top) = shape.GetBounds();
            return new VisioShapeBounds(left, bottom, right, top);
        }

        /// <summary>
        /// Gets aggregate bounds for a shape sequence.
        /// </summary>
        /// <param name="shapes">Shapes to inspect.</param>
        public static VisioShapeBounds GetShapeBounds(this IEnumerable<VisioShape> shapes) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            bool any = false;
            double left = 0;
            double bottom = 0;
            double right = 0;
            double top = 0;

            foreach (VisioShape shape in shapes) {
                VisioShapeBounds bounds = shape.GetShapeBounds();
                if (!any) {
                    left = bounds.Left;
                    bottom = bounds.Bottom;
                    right = bounds.Right;
                    top = bounds.Top;
                    any = true;
                    continue;
                }

                left = Math.Min(left, bounds.Left);
                bottom = Math.Min(bottom, bounds.Bottom);
                right = Math.Max(right, bounds.Right);
                top = Math.Max(top, bounds.Top);
            }

            return any ? new VisioShapeBounds(left, bottom, right, top) : VisioShapeBounds.Empty;
        }

        /// <summary>
        /// Gets bounds of visible page content, including shapes, explicit connector routes, and connector label boxes.
        /// </summary>
        /// <param name="page">Page to inspect.</param>
        /// <param name="includeGroupChildren">Whether nested group children should be included using their stored coordinates.</param>
        /// <param name="includeConnectors">Whether connector waypoints and label boxes should be included.</param>
        public static VisioShapeBounds GetContentBounds(this VisioPage page, bool includeGroupChildren = false, bool includeConnectors = true) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioShapeBounds bounds = includeGroupChildren ? page.AllShapes().GetShapeBounds() : page.Shapes.GetShapeBounds();
            if (!includeConnectors) {
                return bounds;
            }

            foreach (VisioConnector connector in page.Connectors) {
                bounds = Combine(bounds, GetConnectorContentBounds(connector));
            }

            return bounds;
        }

        /// <summary>
        /// Moves top-level page shapes so content starts at the requested margin and optionally resizes the page to fit.
        /// </summary>
        /// <param name="page">Page to update.</param>
        /// <param name="margin">Margin in inches around content.</param>
        /// <param name="resizePage">Whether to resize the page around the content.</param>
        public static VisioPage FitToContent(this VisioPage page, double margin = 0.5D, bool resizePage = true) {
            return FitToContent(page, margin, margin, resizePage);
        }

        /// <summary>
        /// Moves top-level page shapes so content starts at the requested margins and optionally resizes the page to fit.
        /// </summary>
        /// <param name="page">Page to update.</param>
        /// <param name="horizontalMargin">Horizontal margin in inches.</param>
        /// <param name="verticalMargin">Vertical margin in inches.</param>
        /// <param name="resizePage">Whether to resize the page around the content.</param>
        public static VisioPage FitToContent(this VisioPage page, double horizontalMargin, double verticalMargin, bool resizePage = true) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (horizontalMargin < 0) {
                throw new ArgumentOutOfRangeException(nameof(horizontalMargin), "Margin cannot be negative.");
            }

            if (verticalMargin < 0) {
                throw new ArgumentOutOfRangeException(nameof(verticalMargin), "Margin cannot be negative.");
            }

            VisioShapeBounds bounds = page.GetContentBounds();
            if (bounds.IsEmpty) {
                return page;
            }

            double deltaX = horizontalMargin - bounds.Left;
            double deltaY = verticalMargin - bounds.Bottom;
            MoveShapes(page.Shapes, deltaX, deltaY);
            MoveConnectorPageCoordinates(page.Connectors, deltaX, deltaY);

            if (resizePage) {
                page.Width = Math.Max(MinimumPageSize, bounds.Width + horizontalMargin * 2D);
                page.Height = Math.Max(MinimumPageSize, bounds.Height + verticalMargin * 2D);
            }

            return page;
        }

        /// <summary>
        /// Centers top-level page shapes within the current page size.
        /// </summary>
        /// <param name="page">Page to update.</param>
        public static VisioPage CenterContent(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioShapeBounds bounds = page.GetContentBounds();
            if (bounds.IsEmpty) {
                return page;
            }

            MoveShapes(page.Shapes, (page.Width / 2D) - bounds.CenterX, (page.Height / 2D) - bounds.CenterY);
            MoveConnectorPageCoordinates(page.Connectors, (page.Width / 2D) - bounds.CenterX, (page.Height / 2D) - bounds.CenterY);
            return page;
        }

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

        /// <summary>
        /// Resizes a shape to fit its plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="shape">Shape to resize.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses Office default when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting height in inches.</param>
        public static VisioShape ResizeToText(this VisioShape shape, OfficeFontInfo? fontInfo = null, double horizontalPadding = DefaultHorizontalPadding, double verticalPadding = DefaultVerticalPadding, double minimumWidth = 0.5D, double minimumHeight = 0.3D) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            ValidateTextResizeArguments(horizontalPadding, verticalPadding, minimumWidth, minimumHeight);
            (double width, double height) = MeasureTextBox(
                shape.Text,
                fontInfo ?? CreateFontInfo(shape.TextStyle),
                horizontalPadding,
                verticalPadding,
                minimumWidth,
                minimumHeight,
                maximumWidth: null,
                shape.TextStyle);

            shape.Width = width;
            shape.Height = height;
            shape.LocPinX = shape.Width / 2D;
            shape.LocPinY = shape.Height / 2D;
            return shape;
        }

        /// <summary>
        /// Resizes selected shapes to fit their plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="selection">Selection to resize.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses Office default when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting height in inches.</param>
        public static VisioShapeSelection ResizeToText(this VisioShapeSelection selection, OfficeFontInfo? fontInfo = null, double horizontalPadding = DefaultHorizontalPadding, double verticalPadding = DefaultVerticalPadding, double minimumWidth = 0.5D, double minimumHeight = 0.3D) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioShape shape in selection) {
                shape.ResizeToText(fontInfo, horizontalPadding, verticalPadding, minimumWidth, minimumHeight);
            }

            return selection;
        }

        /// <summary>
        /// Resizes a connector label text box to fit its plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="connector">Connector whose label box should be resized.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses connector text style, then Office default, when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting label width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting label height in inches.</param>
        /// <param name="maximumWidth">Optional maximum label width in inches. Text wraps by words when supplied.</param>
        public static VisioConnector ResizeLabelToText(this VisioConnector connector, OfficeFontInfo? fontInfo = null, double horizontalPadding = 0.12D, double verticalPadding = 0.06D, double minimumWidth = 0.45D, double minimumHeight = 0.22D, double? maximumWidth = null) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            ValidateTextResizeArguments(horizontalPadding, verticalPadding, minimumWidth, minimumHeight);
            if (maximumWidth.HasValue && maximumWidth.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumWidth), "Maximum width must be positive.");
            }

            VisioConnectorLabelPlacement placement = connector.LabelPlacement ?? VisioConnectorLabelPlacement.Along(0.5D);
            (double width, double height) = MeasureTextBox(
                connector.Label,
                fontInfo ?? CreateFontInfo(connector.TextStyle),
                horizontalPadding,
                verticalPadding,
                minimumWidth,
                minimumHeight,
                maximumWidth,
                connector.TextStyle);

            placement.Width = width;
            placement.Height = height;
            connector.LabelPlacement = placement;
            return connector;
        }

        /// <summary>
        /// Resizes selected connector label text boxes to fit their plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses connector text style, then Office default, when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting label width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting label height in inches.</param>
        /// <param name="maximumWidth">Optional maximum label width in inches. Text wraps by words when supplied.</param>
        public static VisioConnectorSelection ResizeLabelsToText(this VisioConnectorSelection selection, OfficeFontInfo? fontInfo = null, double horizontalPadding = 0.12D, double verticalPadding = 0.06D, double minimumWidth = 0.45D, double minimumHeight = 0.22D, double? maximumWidth = null) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioConnector connector in selection) {
                connector.ResizeLabelToText(fontInfo, horizontalPadding, verticalPadding, minimumWidth, minimumHeight, maximumWidth);
            }

            return selection;
        }

        /// <summary>
        /// Moves connector label boxes away from page edges, unrelated shapes, and previously placed connector labels.
        /// </summary>
        /// <param name="page">Page whose connector labels should be adjusted.</param>
        /// <param name="step">Search step in page units, expressed in inches.</param>
        /// <param name="maxAttempts">Number of search rings to try around the current label position.</param>
        /// <param name="avoidShapes">Whether labels should avoid unrelated non-container shapes.</param>
        /// <param name="avoidLabels">Whether labels should avoid other connector labels.</param>
        public static VisioPage ResolveConnectorLabelOverlaps(this VisioPage page, double step = 0.18D, int maxAttempts = 12, bool avoidShapes = true, bool avoidLabels = true) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (step <= 0D || double.IsNaN(step) || double.IsInfinity(step)) {
                throw new ArgumentOutOfRangeException(nameof(step), "Step must be a positive finite value.");
            }

            if (maxAttempts < 0) {
                throw new ArgumentOutOfRangeException(nameof(maxAttempts), "Attempt count cannot be negative.");
            }

            IReadOnlyList<VisioShape> shapes = page.Shapes.ToList();
            Dictionary<VisioShape, VisioShapeBounds> shapeBounds = shapes.ToDictionary(shape => shape, shape => shape.GetShapeBounds());
            List<ConnectorLabelBounds> placedLabels = new();

            foreach (VisioConnector connector in page.Connectors) {
                VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
                if (placement == null || string.IsNullOrWhiteSpace(connector.Label)) {
                    continue;
                }

                List<Point> path = BuildConnectorPath(connector);
                if (!TryGetConnectorLabelBounds(connector, path, placement, out VisioShapeBounds currentBounds)) {
                    continue;
                }

                CandidateScore currentScore = ScoreConnectorLabel(page, connector, currentBounds, shapes, shapeBounds, placedLabels, avoidShapes, avoidLabels);
                if (!currentScore.HasConflict) {
                    placedLabels.Add(new ConnectorLabelBounds(connector, currentBounds));
                    continue;
                }

                VisioConnectorLabelPlacement bestPlacement = placement.Clone();
                VisioShapeBounds bestBounds = currentBounds;
                CandidateScore bestScore = currentScore;

                foreach (LabelCandidate candidate in EnumerateLabelCandidates(maxAttempts, step)) {
                    VisioConnectorLabelPlacement candidatePlacement = CreateCandidatePlacement(placement, candidate);
                    if (!TryGetConnectorLabelBounds(connector, path, candidatePlacement, out VisioShapeBounds candidateBounds)) {
                        continue;
                    }

                    CandidateScore candidateScore = ScoreConnectorLabel(page, connector, candidateBounds, shapes, shapeBounds, placedLabels, avoidShapes, avoidLabels);
                    if (candidateScore.IsBetterThan(bestScore)) {
                        bestPlacement = candidatePlacement;
                        bestBounds = candidateBounds;
                        bestScore = candidateScore;
                    }

                    if (!candidateScore.HasConflict) {
                        break;
                    }
                }

                connector.LabelPlacement = bestPlacement;
                placedLabels.Add(new ConnectorLabelBounds(connector, bestBounds));
            }

            return page;
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

        private static string[] SplitLines(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return new[] { string.Empty };
            }

            return text!.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
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

        private static IEnumerable<LabelCandidate> EnumerateLabelCandidates(int maxAttempts, double step) {
            yield return new LabelCandidate(0D, 0D);
            for (int ring = 1; ring <= maxAttempts; ring++) {
                double distance = ring * step;
                yield return new LabelCandidate(0D, distance);
                yield return new LabelCandidate(0D, -distance);
                yield return new LabelCandidate(distance, 0D);
                yield return new LabelCandidate(-distance, 0D);
                yield return new LabelCandidate(distance, distance);
                yield return new LabelCandidate(-distance, distance);
                yield return new LabelCandidate(distance, -distance);
                yield return new LabelCandidate(-distance, -distance);
            }
        }

        private static VisioConnectorLabelPlacement CreateCandidatePlacement(VisioConnectorLabelPlacement source, LabelCandidate candidate) {
            VisioConnectorLabelPlacement placement = source.Clone();
            if (placement.AbsolutePinX.HasValue && placement.AbsolutePinY.HasValue) {
                placement.AbsolutePinX += candidate.OffsetX;
                placement.AbsolutePinY += candidate.OffsetY;
            } else {
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
            bool avoidShapes,
            bool avoidLabels) {
            double pagePenalty = OutsidePageAmount(labelBounds, page);
            double shapeOverlap = 0D;
            if (avoidShapes) {
                foreach (VisioShape shape in shapes) {
                    if (ReferenceEquals(shape, connector.From) || ReferenceEquals(shape, connector.To)) {
                        continue;
                    }

                    if (shape.IsContainer || shape.IsBackgroundSurface) {
                        continue;
                    }

                    VisioShapeBounds bounds = shapeBounds[shape];
                    if (Contains(bounds, labelBounds)) {
                        continue;
                    }

                    shapeOverlap += OverlapArea(labelBounds, bounds);
                }
            }

            double labelOverlap = 0D;
            if (avoidLabels) {
                foreach (ConnectorLabelBounds placedLabel in placedLabels) {
                    labelOverlap += OverlapArea(labelBounds, placedLabel.Bounds);
                }
            }

            return new CandidateScore(pagePenalty, shapeOverlap, labelOverlap);
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

        private static void ValidateTextResizeArguments(double horizontalPadding, double verticalPadding, double minimumWidth, double minimumHeight) {
            if (horizontalPadding < 0) {
                throw new ArgumentOutOfRangeException(nameof(horizontalPadding), "Padding cannot be negative.");
            }

            if (verticalPadding < 0) {
                throw new ArgumentOutOfRangeException(nameof(verticalPadding), "Padding cannot be negative.");
            }

            if (minimumWidth < 0) {
                throw new ArgumentOutOfRangeException(nameof(minimumWidth), "Minimum size cannot be negative.");
            }

            if (minimumHeight < 0) {
                throw new ArgumentOutOfRangeException(nameof(minimumHeight), "Minimum size cannot be negative.");
            }
        }

        private static (double Width, double Height) MeasureTextBox(
            string? text,
            OfficeFontInfo fontInfo,
            double horizontalPadding,
            double verticalPadding,
            double minimumWidth,
            double minimumHeight,
            double? maximumWidth,
            VisioTextStyle? textStyle) {
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(fontInfo);
            OfficeTextMeasurementStyle style = measurer.CreateStyle(fontInfo);
            double horizontalMargins = (textStyle?.LeftMargin ?? 0D) + (textStyle?.RightMargin ?? 0D);
            double verticalMargins = (textStyle?.TopMargin ?? 0D) + (textStyle?.BottomMargin ?? 0D);
            double fixedWidth = horizontalPadding * 2D + horizontalMargins;
            double fixedHeight = verticalPadding * 2D + verticalMargins;
            double? maximumContentWidthPixels = null;
            if (maximumWidth.HasValue) {
                double contentWidth = Math.Max(0.01D, maximumWidth.Value - fixedWidth);
                maximumContentWidthPixels = contentWidth * style.Dpi;
            }

            string[] lines = WrapLines(SplitLines(text), measurer, style, maximumContentWidthPixels);
            double maxWidthPixels = 0;
            foreach (string line in lines) {
                maxWidthPixels = Math.Max(maxWidthPixels, measurer.MeasureWidth(line, style));
            }

            double lineHeightPixels = measurer.MeasureLineHeight(style);
            double measuredWidth = maxWidthPixels / style.Dpi + fixedWidth;
            double measuredHeight = (lineHeightPixels * Math.Max(1, lines.Length)) / style.Dpi + fixedHeight;
            double width = Math.Max(minimumWidth, measuredWidth);
            if (maximumWidth.HasValue) {
                width = Math.Min(Math.Max(minimumWidth, maximumWidth.Value), width);
            }

            return (width, Math.Max(minimumHeight, measuredHeight));
        }

        private static OfficeFontInfo CreateFontInfo(VisioTextStyle? textStyle) {
            if (textStyle == null) {
                return OfficeFontInfo.Default;
            }

            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (textStyle.Bold == true) {
                style |= OfficeFontStyle.Bold;
            }

            if (textStyle.Italic == true) {
                style |= OfficeFontStyle.Italic;
            }

            if (textStyle.Underline == true) {
                style |= OfficeFontStyle.Underline;
            }

            return new OfficeFontInfo(
                string.IsNullOrWhiteSpace(textStyle.FontFamily) ? OfficeFontInfo.Default.FamilyName : textStyle.FontFamily,
                textStyle.Size ?? OfficeFontInfo.Default.Size,
                style);
        }

        private static string[] WrapLines(string[] sourceLines, OfficeTextMeasurer measurer, OfficeTextMeasurementStyle style, double? maximumContentWidthPixels) {
            if (!maximumContentWidthPixels.HasValue) {
                return sourceLines;
            }

            List<string> wrapped = new();
            foreach (string sourceLine in sourceLines) {
                WrapLine(sourceLine, measurer, style, maximumContentWidthPixels.Value, wrapped);
            }

            return wrapped.Count == 0 ? new[] { string.Empty } : wrapped.ToArray();
        }

        private static void WrapLine(string sourceLine, OfficeTextMeasurer measurer, OfficeTextMeasurementStyle style, double maximumContentWidthPixels, IList<string> destination) {
            string[] words = sourceLine.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) {
                destination.Add(string.Empty);
                return;
            }

            string current = string.Empty;
            foreach (string word in words) {
                string candidate = current.Length == 0 ? word : current + " " + word;
                if (current.Length > 0 && measurer.MeasureWidth(candidate, style) > maximumContentWidthPixels) {
                    destination.Add(current);
                    current = word;
                } else {
                    current = candidate;
                }
            }

            destination.Add(current);
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

        private readonly struct LabelCandidate {
            public LabelCandidate(double offsetX, double offsetY) {
                OffsetX = offsetX;
                OffsetY = offsetY;
            }

            public double OffsetX { get; }

            public double OffsetY { get; }
        }

        private readonly struct ConnectorLabelBounds {
            public ConnectorLabelBounds(VisioConnector connector, VisioShapeBounds bounds) {
                Connector = connector;
                Bounds = bounds;
            }

            public VisioConnector Connector { get; }

            public VisioShapeBounds Bounds { get; }
        }

        private readonly struct CandidateScore {
            public CandidateScore(double pagePenalty, double shapeOverlap, double labelOverlap) {
                PagePenalty = pagePenalty;
                ShapeOverlap = shapeOverlap;
                LabelOverlap = labelOverlap;
            }

            public double PagePenalty { get; }

            public double ShapeOverlap { get; }

            public double LabelOverlap { get; }

            public bool HasConflict => PagePenalty > 1e-9 || ShapeOverlap > 1e-9 || LabelOverlap > 1e-9;

            public bool IsBetterThan(CandidateScore other) {
                if (PagePenalty < other.PagePenalty - 1e-9) {
                    return true;
                }

                if (PagePenalty > other.PagePenalty + 1e-9) {
                    return false;
                }

                if (ShapeOverlap < other.ShapeOverlap - 1e-9) {
                    return true;
                }

                if (ShapeOverlap > other.ShapeOverlap + 1e-9) {
                    return false;
                }

                return LabelOverlap < other.LabelOverlap - 1e-9;
            }
        }
    }
}
