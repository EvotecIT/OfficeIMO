using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// High-level diagram cleanup helpers that combine common layout, label, and page-fit passes.
    /// </summary>
    public static class VisioDiagramPolishExtensions {
        /// <summary>
        /// Applies a deterministic cleanup pass to every foreground page in the document.
        /// </summary>
        /// <param name="document">Document whose pages should be polished.</param>
        /// <param name="options">Optional polish settings.</param>
        public static VisioDocument PolishDiagrams(this VisioDocument document, VisioDiagramPolishOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            VisioDiagramPolishOptions resolvedOptions = options ?? new VisioDiagramPolishOptions();
            foreach (VisioPage page in document.Pages) {
                if (page.IsBackground && !resolvedOptions.IncludeBackgroundPages) {
                    continue;
                }

                page.PolishDiagram(resolvedOptions);
            }

            return document;
        }

        /// <summary>
        /// Applies a deterministic cleanup pass to a page: optional text fitting, connector label fitting, label collision cleanup, and page fitting.
        /// </summary>
        /// <param name="page">Page to polish.</param>
        /// <param name="options">Optional polish settings.</param>
        public static VisioPage PolishDiagram(this VisioPage page, VisioDiagramPolishOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioDiagramPolishOptions resolvedOptions = options ?? new VisioDiagramPolishOptions();
            ValidateOptions(resolvedOptions);

            if (resolvedOptions.ResizeShapesToText) {
                foreach (VisioShape shape in page.Shapes) {
                    if (shape.IsContainer || shape.IsBackgroundSurface) {
                        continue;
                    }

                    shape.ResizeToText(
                        resolvedOptions.ShapeFontInfo,
                        resolvedOptions.ShapeHorizontalPadding,
                        resolvedOptions.ShapeVerticalPadding,
                        resolvedOptions.MinimumShapeWidth,
                        resolvedOptions.MinimumShapeHeight);
                }
            }

            if (resolvedOptions.ResizeConnectorLabelsToText) {
                foreach (VisioConnector connector in page.Connectors) {
                    if (string.IsNullOrWhiteSpace(connector.Label)) {
                        continue;
                    }

                    connector.ResizeLabelToText(
                        resolvedOptions.ConnectorLabelFontInfo,
                        resolvedOptions.ConnectorLabelHorizontalPadding,
                        resolvedOptions.ConnectorLabelVerticalPadding,
                        resolvedOptions.MinimumConnectorLabelWidth,
                        resolvedOptions.MinimumConnectorLabelHeight,
                        resolvedOptions.MaximumConnectorLabelWidth);
                }
            }

            if (resolvedOptions.ResolveShapeOverlaps) {
                page.ResolveShapeOverlaps(
                    resolvedOptions.ShapeOverlapStep,
                    resolvedOptions.ShapeOverlapMaxAttempts,
                    resolvedOptions.IncludeContainersInShapeOverlapResolution);
            }

            if (resolvedOptions.ResolveConnectorShapeIntersections) {
                page.RouteConnectorsOrthogonalAroundShapes(
                    resolvedOptions.ConnectorRoutingObstaclePadding,
                    resolvedOptions.ConnectorRoutingMaxLanes);
            }

            if (resolvedOptions.ResolveConnectorLabelOverlaps) {
                page.ResolveConnectorLabelOverlaps(
                    resolvedOptions.ConnectorLabelStep,
                    resolvedOptions.ConnectorLabelMaxAttempts,
                    resolvedOptions.AvoidConnectorLabelShapeOverlaps,
                    resolvedOptions.AvoidConnectorLabelOverlaps);
            }

            if (resolvedOptions.FitToContent) {
                page.FitToContent(resolvedOptions.FitHorizontalMargin, resolvedOptions.FitVerticalMargin, resolvedOptions.ResizePage);
            }

            return page;
        }

        private static void ValidateOptions(VisioDiagramPolishOptions options) {
            if (options.FitHorizontalMargin < 0D) {
                throw new ArgumentOutOfRangeException(nameof(options), "Fit margins cannot be negative.");
            }

            if (options.FitVerticalMargin < 0D) {
                throw new ArgumentOutOfRangeException(nameof(options), "Fit margins cannot be negative.");
            }

            if (options.ConnectorLabelStep <= 0D || double.IsNaN(options.ConnectorLabelStep) || double.IsInfinity(options.ConnectorLabelStep)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Connector label step must be a positive finite value.");
            }

            if (options.ConnectorLabelMaxAttempts < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Connector label attempt count cannot be negative.");
            }

            if (options.ConnectorRoutingObstaclePadding < 0D ||
                double.IsNaN(options.ConnectorRoutingObstaclePadding) ||
                double.IsInfinity(options.ConnectorRoutingObstaclePadding)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Connector routing obstacle padding must be a non-negative finite value.");
            }

            if (options.ConnectorRoutingMaxLanes < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Connector routing lane count cannot be negative.");
            }

            if (options.ShapeOverlapStep <= 0D || double.IsNaN(options.ShapeOverlapStep) || double.IsInfinity(options.ShapeOverlapStep)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Shape overlap step must be a positive finite value.");
            }

            if (options.ShapeOverlapMaxAttempts < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Shape overlap attempt count cannot be negative.");
            }
        }
    }
}
