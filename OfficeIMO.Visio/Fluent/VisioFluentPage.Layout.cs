using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Aligns existing shapes by id using the current selection bounds.
        /// </summary>
        /// <param name="alignment">Horizontal alignment to apply.</param>
        /// <param name="shapeIds">Shape ids to align.</param>
        public VisioFluentPage AlignShapes(VisioHorizontalAlignment alignment, IEnumerable<string> shapeIds) {
            ResolveShapeSelection(shapeIds).Align(alignment);
            return this;
        }

        /// <summary>
        /// Aligns existing shapes by id using the current selection bounds.
        /// </summary>
        /// <param name="alignment">Vertical alignment to apply.</param>
        /// <param name="shapeIds">Shape ids to align.</param>
        public VisioFluentPage AlignShapes(VisioVerticalAlignment alignment, IEnumerable<string> shapeIds) {
            ResolveShapeSelection(shapeIds).Align(alignment);
            return this;
        }

        /// <summary>
        /// Distributes existing shapes by center point along the requested axis.
        /// </summary>
        /// <param name="axis">Distribution axis.</param>
        /// <param name="shapeIds">Shape ids to distribute.</param>
        public VisioFluentPage DistributeShapes(VisioDistributionAxis axis, IEnumerable<string> shapeIds) {
            ResolveShapeSelection(shapeIds).Distribute(axis);
            return this;
        }

        /// <summary>
        /// Relays out existing shapes by id into a deterministic grid and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="shapeIds">Shape ids to relayout.</param>
        /// <param name="options">Optional layout settings.</param>
        public VisioFluentPage RelayoutShapesAsGrid(IEnumerable<string> shapeIds, VisioSelectionLayoutOptions? options = null) {
            ResolveShapeSelection(shapeIds).RelayoutAsGrid(options);
            return this;
        }

        /// <summary>
        /// Relays out existing shapes by id into a deterministic grid using inline option configuration.
        /// </summary>
        /// <param name="shapeIds">Shape ids to relayout.</param>
        /// <param name="configureOptions">Option callback.</param>
        public VisioFluentPage RelayoutShapesAsGrid(IEnumerable<string> shapeIds, Action<VisioSelectionLayoutOptions> configureOptions) {
            if (configureOptions == null) {
                throw new ArgumentNullException(nameof(configureOptions));
            }

            VisioSelectionLayoutOptions options = new();
            configureOptions(options);
            return RelayoutShapesAsGrid(shapeIds, options);
        }

        /// <summary>
        /// Relays out existing shapes by id as a horizontal row and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="shapeIds">Shape ids to relayout.</param>
        /// <param name="spacing">Horizontal spacing between shapes in inches.</param>
        /// <param name="routeInternalConnectors">Whether connectors whose endpoints are both selected should be rerouted orthogonally.</param>
        public VisioFluentPage RelayoutShapesAsHorizontalStack(IEnumerable<string> shapeIds, double spacing = 0.5D, bool routeInternalConnectors = true) {
            ResolveShapeSelection(shapeIds).RelayoutAsHorizontalStack(spacing, routeInternalConnectors);
            return this;
        }

        /// <summary>
        /// Relays out existing shapes by id as a vertical stack and optionally reroutes internal connectors.
        /// </summary>
        /// <param name="shapeIds">Shape ids to relayout.</param>
        /// <param name="spacing">Vertical spacing between shapes in inches.</param>
        /// <param name="routeInternalConnectors">Whether connectors whose endpoints are both selected should be rerouted orthogonally.</param>
        public VisioFluentPage RelayoutShapesAsVerticalStack(IEnumerable<string> shapeIds, double spacing = 0.5D, bool routeInternalConnectors = true) {
            ResolveShapeSelection(shapeIds).RelayoutAsVerticalStack(spacing, routeInternalConnectors);
            return this;
        }

        /// <summary>
        /// Relays out all shapes reachable from a starting shape through connectors.
        /// </summary>
        /// <param name="shapeId">Starting shape id.</param>
        /// <param name="options">Optional layout settings.</param>
        /// <param name="includeStart">Whether the starting shape should be included in the relayout.</param>
        public VisioFluentPage RelayoutConnectedComponentAsGrid(string shapeId, VisioSelectionLayoutOptions? options = null, bool includeStart = true) {
            Page.SelectConnectedComponent(ResolveShape(shapeId), includeStart).RelayoutAsGrid(options);
            return this;
        }

        /// <summary>
        /// Relays out the typed members of a Visio-native container and optionally refits the container around them.
        /// </summary>
        /// <param name="containerId">Container shape id.</param>
        /// <param name="layoutOptions">Optional member layout settings.</param>
        /// <param name="refitContainer">Whether the container should be resized after member relayout.</param>
        /// <param name="refitOptions">Optional container refit settings.</param>
        public VisioFluentPage RelayoutContainerMembers(
            string containerId,
            VisioSelectionLayoutOptions? layoutOptions = null,
            bool refitContainer = true,
            VisioContainerOptions? refitOptions = null) {
            VisioShape container = ResolveShape(containerId);
            VisioShapeSelection members = new(Page.GetContainerMembers(container), Page);
            members.RelayoutAsGrid(layoutOptions);

            if (refitContainer) {
                Page.RefitContainer(container, refitOptions);
            }

            return this;
        }

        /// <summary>
        /// Relays out the typed members of a Visio-native container using inline layout option configuration.
        /// </summary>
        /// <param name="containerId">Container shape id.</param>
        /// <param name="configureLayoutOptions">Layout option callback.</param>
        /// <param name="refitContainer">Whether the container should be resized after member relayout.</param>
        /// <param name="refitOptions">Optional container refit settings.</param>
        public VisioFluentPage RelayoutContainerMembers(
            string containerId,
            Action<VisioSelectionLayoutOptions> configureLayoutOptions,
            bool refitContainer = true,
            VisioContainerOptions? refitOptions = null) {
            if (configureLayoutOptions == null) {
                throw new ArgumentNullException(nameof(configureLayoutOptions));
            }

            VisioSelectionLayoutOptions options = new();
            configureLayoutOptions(options);
            return RelayoutContainerMembers(containerId, options, refitContainer, refitOptions);
        }

        private VisioShapeSelection ResolveShapeSelection(IEnumerable<string> shapeIds) {
            return new VisioShapeSelection(ResolveShapes(shapeIds, nameof(shapeIds)), Page);
        }
    }
}
