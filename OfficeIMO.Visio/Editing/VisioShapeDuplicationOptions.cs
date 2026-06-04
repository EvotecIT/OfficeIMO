using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options controlling how shapes and internal connectors are duplicated on a Visio page.
    /// </summary>
    public sealed class VisioShapeDuplicationOptions {
        /// <summary>
        /// Horizontal offset in inches for duplicated top-level shapes and page-coordinate connector points.
        /// </summary>
        public double OffsetX { get; set; } = 0.35D;

        /// <summary>
        /// Vertical offset in inches for duplicated top-level shapes and page-coordinate connector points.
        /// </summary>
        public double OffsetY { get; set; } = -0.35D;

        /// <summary>
        /// Whether connectors between duplicated shapes should also be copied.
        /// </summary>
        public bool IncludeInternalConnectors { get; set; } = true;

        /// <summary>
        /// Optional suffix applied to duplicated shape identifiers. Collisions receive a numeric suffix automatically.
        /// </summary>
        public string? IdSuffix { get; set; }

        /// <summary>
        /// Optional suffix applied to duplicated connector identifiers. When omitted, <see cref="IdSuffix"/> is used.
        /// </summary>
        public string? ConnectorIdSuffix { get; set; }

        /// <summary>
        /// Optional factory that returns the preferred identifier for each duplicated shape.
        /// </summary>
        public Func<VisioShape, string?>? ShapeIdFactory { get; set; }

        /// <summary>
        /// Optional factory that returns the preferred identifier for each duplicated internal connector.
        /// </summary>
        public Func<VisioConnector, string?>? ConnectorIdFactory { get; set; }
    }
}
