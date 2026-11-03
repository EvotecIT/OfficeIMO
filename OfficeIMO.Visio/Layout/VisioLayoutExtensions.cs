using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Layout and geometry helpers for Visio pages, shapes, and selections.
    /// </summary>
    public static partial class VisioLayoutExtensions {
        private const double MinimumPageSize = 0.1D;
        private const double DefaultHorizontalPadding = 0.25D;
        private const double DefaultVerticalPadding = 0.14D;

        /// <summary>
        /// Gets the page-space bounds for a shape.
        /// </summary>
        /// <param name="shape">Shape to inspect.</param>
        public static VisioShapeBounds GetShapeBounds(this VisioShape shape) {
            ArgumentNullException.ThrowIfNull(shape);

            (double left, double bottom, double right, double top) = shape.GetBoundsInches();
            return new VisioShapeBounds(left, bottom, right, top);
        }

        /// <summary>
        /// Gets rotation-aware bounds in page coordinates, including every parent-group
        /// transform for a nested shape.
        /// </summary>
        public static VisioShapeBounds GetPageShapeBounds(this VisioShape shape) {
            return shape.GetShapeBounds();
        }

        /// <summary>
        /// Gets aggregate bounds for a shape sequence.
        /// </summary>
        /// <param name="shapes">Shapes to inspect.</param>
        public static VisioShapeBounds GetShapeBounds(this IEnumerable<VisioShape> shapes) {
            ArgumentNullException.ThrowIfNull(shapes);

            bool any = false;
            double left = 0;
            double bottom = 0;
            double right = 0;
            double top = 0;

            foreach (VisioShape shape in shapes) {
                VisioShapeBounds bounds = shape.GetShapeBounds();
                if (!any) {
                    left = bounds.LeftInches;
                    bottom = bounds.BottomInches;
                    right = bounds.RightInches;
                    top = bounds.TopInches;
                    any = true;
                    continue;
                }

                left = Math.Min(left, bounds.LeftInches);
                bottom = Math.Min(bottom, bounds.BottomInches);
                right = Math.Max(right, bounds.RightInches);
                top = Math.Max(top, bounds.TopInches);
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
            ArgumentNullException.ThrowIfNull(page);

            VisioShapeBounds bounds = page.Shapes.GetShapeBounds();
            if (includeGroupChildren) {
                foreach (VisioShape child in page.AllShapes().Where(shape => shape.Parent != null))
                    bounds = Combine(bounds, child.GetPageShapeBounds());
            }
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
        /// Fits an imported page using explicit group-child, connector-route, translation,
        /// and resize policies. Rotation-aware shape bounds and connector label boxes are used.
        /// </summary>
        public static VisioPage FitToContent(this VisioPage page,
            VisioFitToContentOptions options) {
            ArgumentNullException.ThrowIfNull(page);
            ArgumentNullException.ThrowIfNull(options);
            if (options.HorizontalMargin < 0D || double.IsNaN(options.HorizontalMargin) || double.IsInfinity(options.HorizontalMargin))
                throw new ArgumentOutOfRangeException(nameof(options), "Horizontal margin must be non-negative and finite.");
            if (options.VerticalMargin < 0D || double.IsNaN(options.VerticalMargin) || double.IsInfinity(options.VerticalMargin))
                throw new ArgumentOutOfRangeException(nameof(options), "Vertical margin must be non-negative and finite.");
            VisioShapeBounds bounds = page.GetContentBounds(options.IncludeGroupChildren,
                options.IncludeConnectors);
            if (bounds.IsEmpty) return page;
            if (options.MoveContent) {
                double deltaX = options.HorizontalMargin - bounds.LeftInches;
                double deltaY = options.VerticalMargin - bounds.BottomInches;
                MoveShapes(page.Shapes, deltaX, deltaY);
                MoveConnectorPageCoordinates(page.Connectors, deltaX, deltaY);
            }
            if (options.ResizePage) {
                page.WidthInches = Math.Max(MinimumPageSize, options.MoveContent
                    ? bounds.WidthInches + options.HorizontalMargin * 2D
                    : Math.Max(0D, bounds.RightInches) + options.HorizontalMargin);
                page.HeightInches = Math.Max(MinimumPageSize, options.MoveContent
                    ? bounds.HeightInches + options.VerticalMargin * 2D
                    : Math.Max(0D, bounds.TopInches) + options.VerticalMargin);
            }
            return page;
        }

        /// <summary>
        /// Moves top-level page shapes so content starts at the requested margins and optionally resizes the page to fit.
        /// </summary>
        /// <param name="page">Page to update.</param>
        /// <param name="horizontalMargin">Horizontal margin in inches.</param>
        /// <param name="verticalMargin">Vertical margin in inches.</param>
        /// <param name="resizePage">Whether to resize the page around the content.</param>
        public static VisioPage FitToContent(this VisioPage page, double horizontalMargin, double verticalMargin, bool resizePage = true) {
            ArgumentNullException.ThrowIfNull(page);

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

            double deltaX = horizontalMargin - bounds.LeftInches;
            double deltaY = verticalMargin - bounds.BottomInches;
            MoveShapes(page.Shapes, deltaX, deltaY);
            MoveConnectorPageCoordinates(page.Connectors, deltaX, deltaY);

            if (resizePage) {
                page.WidthInches = Math.Max(MinimumPageSize, bounds.WidthInches + horizontalMargin * 2D);
                page.HeightInches = Math.Max(MinimumPageSize, bounds.HeightInches + verticalMargin * 2D);
            }

            return page;
        }

        /// <summary>
        /// Centers top-level page shapes within the current page size.
        /// </summary>
        /// <param name="page">Page to update.</param>
        public static VisioPage CenterContent(this VisioPage page) {
            ArgumentNullException.ThrowIfNull(page);

            VisioShapeBounds bounds = page.GetContentBounds();
            if (bounds.IsEmpty) {
                return page;
            }

            MoveShapes(page.Shapes, (page.WidthInches / 2D) - bounds.CenterXInches, (page.HeightInches / 2D) - bounds.CenterYInches);
            MoveConnectorPageCoordinates(page.Connectors, (page.WidthInches / 2D) - bounds.CenterXInches, (page.HeightInches / 2D) - bounds.CenterYInches);
            return page;
        }

        /// <summary>
        /// Moves overlapping top-level shapes apart using a deterministic nearest-open-position search.
        /// </summary>
        /// <param name="page">Page to update.</param>
        /// <param name="step">Search step in inches.</param>
        /// <param name="maxAttempts">Number of search rings to try around each overlapping shape.</param>
        /// <param name="includeContainers">Whether container and background surface shapes should be moved and treated as obstacles.</param>
        public static VisioPage ResolveShapeOverlaps(this VisioPage page, double step = 0.25D, int maxAttempts = 24, bool includeContainers = false) {
            ArgumentNullException.ThrowIfNull(page);

            if (step <= 0D || double.IsNaN(step) || double.IsInfinity(step)) {
                throw new ArgumentOutOfRangeException(nameof(step), "Step must be a positive finite value.");
            }

            if (maxAttempts < 0) {
                throw new ArgumentOutOfRangeException(nameof(maxAttempts), "Attempt count cannot be negative.");
            }

            List<VisioShape> shapes = page.Shapes
                .Where(shape => includeContainers || (!shape.IsContainer && !shape.IsBackgroundSurface))
                .ToList();
            if (shapes.Count < 2) {
                return page;
            }

            for (int index = 1; index < shapes.Count; index++) {
                VisioShape shape = shapes[index];
                double initialOverlap = GetTotalShapeOverlap(shape, shapes);
                if (initialOverlap <= 1e-9) {
                    continue;
                }

                double originalX = shape.PinXInches;
                double originalY = shape.PinYInches;
                double bestX = originalX;
                double bestY = originalY;
                double bestOverlap = initialOverlap;

                foreach (ShapeCandidate candidate in EnumerateShapeCandidates(maxAttempts, step)) {
                    if (Math.Abs(candidate.OffsetX) < 1e-9 && Math.Abs(candidate.OffsetY) < 1e-9) {
                        continue;
                    }

                    shape.PinXInches = originalX + candidate.OffsetX;
                    shape.PinYInches = originalY + candidate.OffsetY;
                    double overlap = GetTotalShapeOverlap(shape, shapes);
                    if (overlap <= 1e-9) {
                        bestX = shape.PinXInches;
                        bestY = shape.PinYInches;
                        bestOverlap = overlap;
                        break;
                    }

                    if (overlap < bestOverlap - 1e-9) {
                        bestX = shape.PinXInches;
                        bestY = shape.PinYInches;
                        bestOverlap = overlap;
                    }
                }

                shape.PinXInches = bestX;
                shape.PinYInches = bestY;
            }

            return page;
        }
    }
}
