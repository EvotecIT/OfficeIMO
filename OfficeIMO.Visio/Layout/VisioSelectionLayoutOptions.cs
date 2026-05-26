namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for deterministic selection relayout.
    /// </summary>
    public sealed class VisioSelectionLayoutOptions {
        /// <summary>
        /// Number of columns in the generated grid. When unset or zero, OfficeIMO uses a near-square grid.
        /// </summary>
        public int? Columns { get; set; }

        /// <summary>
        /// Horizontal spacing between grid columns in inches.
        /// </summary>
        public double HorizontalSpacing { get; set; } = 0.5D;

        /// <summary>
        /// Vertical spacing between grid rows in inches.
        /// </summary>
        public double VerticalSpacing { get; set; } = 0.5D;

        /// <summary>
        /// Whether relayout should start at the current selection's top-left bounds.
        /// When false, the first shape keeps its current center.
        /// </summary>
        public bool PreserveTopLeft { get; set; } = true;

        /// <summary>
        /// Order used before assigning shapes to grid cells.
        /// </summary>
        public VisioSelectionLayoutOrder Order { get; set; } = VisioSelectionLayoutOrder.SelectionOrder;

        /// <summary>
        /// Whether connectors whose endpoints are both inside the selection should be rerouted orthogonally.
        /// Requires a page-backed selection.
        /// </summary>
        public bool RouteInternalConnectors { get; set; } = true;

        /// <summary>
        /// Orthogonal routing style used for internal connectors when <see cref="RouteInternalConnectors"/> is enabled.
        /// </summary>
        public VisioConnectorRouteStyle ConnectorRouteStyle { get; set; } = VisioConnectorRouteStyle.Auto;
    }
}
