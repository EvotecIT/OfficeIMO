namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for deterministic connector routing around diagram obstacles.
    /// </summary>
    public sealed class VisioConnectorRoutingOptions {
        /// <summary>Padding added around each obstacle while testing route intersections.</summary>
        public double Padding { get; set; } = 0.15D;

        /// <summary>Number of positive and negative routing lanes to try on each axis.</summary>
        public int MaxLanes { get; set; } = 12;

        /// <summary>Whether Visio container shapes should be treated as route obstacles unless they contain the connector endpoints.</summary>
        public bool IncludeContainers { get; set; }

        /// <summary>Whether background surfaces such as zones, subnets, and trust boundaries should be treated as route obstacles unless they contain the connector endpoints.</summary>
        public bool IncludeBackgroundSurfaces { get; set; }

        /// <summary>Whether generated adornments such as zone captions should be treated as route obstacles.</summary>
        public bool IncludeDiagramAdornments { get; set; }

        /// <summary>
        /// Creates a detached copy of the options.
        /// </summary>
        public VisioConnectorRoutingOptions Clone() {
            return new VisioConnectorRoutingOptions {
                Padding = Padding,
                MaxLanes = MaxLanes,
                IncludeContainers = IncludeContainers,
                IncludeBackgroundSurfaces = IncludeBackgroundSurfaces,
                IncludeDiagramAdornments = IncludeDiagramAdornments
            };
        }
    }
}
