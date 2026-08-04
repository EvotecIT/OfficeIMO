namespace OfficeIMO.Visio {
    /// <summary>Options for deterministic topology-aware relayout of an imported page.</summary>
    public sealed class VisioWholeDiagramRelayoutOptions {
        /// <summary>Horizontal distance between topology layers, in inches.</summary>
        public double LayerSpacing { get; set; } = 2.0D;

        /// <summary>Vertical distance between shapes in one layer, in inches.</summary>
        public double NodeSpacing { get; set; } = 0.65D;

        /// <summary>Whether the primary flow should run from left to right.</summary>
        public bool LeftToRight { get; set; } = true;

        /// <summary>Whether containers and background surfaces may be moved.</summary>
        public bool IncludeContainers { get; set; }

        /// <summary>Whether grouped shapes may be moved as top-level nodes.</summary>
        public bool IncludeGroups { get; set; } = true;

        /// <summary>Whether connectors between moved shapes should be rerouted.</summary>
        public bool RouteConnectors { get; set; } = true;

        /// <summary>Whether the existing polish and resize-to-content pass runs afterward.</summary>
        public bool PolishAfterLayout { get; set; } = true;

        /// <summary>Polish options used after layout.</summary>
        public VisioDiagramPolishOptions PolishOptions { get; set; } = new VisioDiagramPolishOptions {
            ResolveShapeOverlaps = true,
            ResolveConnectorShapeIntersections = true,
            ConnectorRoutingAvoidContainers = true,
            ConnectorRoutingAvoidBackgroundSurfaces = true,
            ConnectorRoutingAvoidDiagramAdornments = true
        };
    }
}
