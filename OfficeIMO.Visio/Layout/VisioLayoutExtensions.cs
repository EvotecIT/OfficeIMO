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
        /// Moves overlapping top-level shapes apart using a deterministic nearest-open-position search.
        /// </summary>
        /// <param name="page">Page to update.</param>
        /// <param name="step">Search step in inches.</param>
        /// <param name="maxAttempts">Number of search rings to try around each overlapping shape.</param>
        /// <param name="includeContainers">Whether container and background surface shapes should be moved and treated as obstacles.</param>
        public static VisioPage ResolveShapeOverlaps(this VisioPage page, double step = 0.25D, int maxAttempts = 24, bool includeContainers = false) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

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

                double originalX = shape.PinX;
                double originalY = shape.PinY;
                double bestX = originalX;
                double bestY = originalY;
                double bestOverlap = initialOverlap;

                foreach (ShapeCandidate candidate in EnumerateShapeCandidates(maxAttempts, step)) {
                    if (Math.Abs(candidate.OffsetX) < 1e-9 && Math.Abs(candidate.OffsetY) < 1e-9) {
                        continue;
                    }

                    shape.PinX = originalX + candidate.OffsetX;
                    shape.PinY = originalY + candidate.OffsetY;
                    double overlap = GetTotalShapeOverlap(shape, shapes);
                    if (overlap <= 1e-9) {
                        bestX = shape.PinX;
                        bestY = shape.PinY;
                        bestOverlap = overlap;
                        break;
                    }

                    if (overlap < bestOverlap - 1e-9) {
                        bestX = shape.PinX;
                        bestY = shape.PinY;
                        bestOverlap = overlap;
                    }
                }

                shape.PinX = bestX;
                shape.PinY = bestY;
            }

            return page;
        }
    }
}
