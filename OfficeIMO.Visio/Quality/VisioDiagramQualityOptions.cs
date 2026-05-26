namespace OfficeIMO.Visio {
    /// <summary>
    /// Options controlling dependency-free diagram visual quality analysis.
    /// </summary>
    public sealed class VisioDiagramQualityOptions {
        /// <summary>Whether to include nested group children when analyzing shape bounds.</summary>
        public bool IncludeGroupChildren { get; set; }

        /// <summary>Whether to report shapes outside the page bounds.</summary>
        public bool CheckPageBounds { get; set; } = true;

        /// <summary>Whether to report overlapping shape bounds.</summary>
        public bool CheckShapeOverlaps { get; set; } = true;

        /// <summary>Whether to ignore overlaps where one shape fully contains another, such as regions or containers.</summary>
        public bool IgnoreContainingShapeOverlaps { get; set; } = true;

        /// <summary>Whether to report explicit connector routes crossing unrelated shapes.</summary>
        public bool CheckConnectorShapeIntersections { get; set; } = true;

        /// <summary>Whether to report connector label placement outside the page.</summary>
        public bool CheckConnectorLabels { get; set; } = true;

        /// <summary>Whether to report connector labels overlapping unrelated shape bounds.</summary>
        public bool CheckConnectorLabelShapeOverlaps { get; set; } = true;

        /// <summary>Whether to report connector labels overlapping each other.</summary>
        public bool CheckConnectorLabelOverlaps { get; set; } = true;

        /// <summary>Whether every connector is expected to have a text label.</summary>
        public bool RequireConnectorLabels { get; set; }

        /// <summary>Allowed distance beyond the page edge before a shape or label is reported.</summary>
        public double PageBoundsTolerance { get; set; } = 0.01D;

        /// <summary>Minimum overlap ratio before overlapping shapes are reported.</summary>
        public double MinimumShapeOverlapRatio { get; set; } = 0.05D;

        /// <summary>Minimum overlap ratio before connector label overlaps are reported.</summary>
        public double MinimumConnectorLabelOverlapRatio { get; set; } = 0.05D;

        /// <summary>Creates a detached copy of these options.</summary>
        public VisioDiagramQualityOptions Clone() {
            return new VisioDiagramQualityOptions {
                IncludeGroupChildren = IncludeGroupChildren,
                CheckPageBounds = CheckPageBounds,
                CheckShapeOverlaps = CheckShapeOverlaps,
                IgnoreContainingShapeOverlaps = IgnoreContainingShapeOverlaps,
                CheckConnectorShapeIntersections = CheckConnectorShapeIntersections,
                CheckConnectorLabels = CheckConnectorLabels,
                CheckConnectorLabelShapeOverlaps = CheckConnectorLabelShapeOverlaps,
                CheckConnectorLabelOverlaps = CheckConnectorLabelOverlaps,
                RequireConnectorLabels = RequireConnectorLabels,
                PageBoundsTolerance = PageBoundsTolerance,
                MinimumShapeOverlapRatio = MinimumShapeOverlapRatio,
                MinimumConnectorLabelOverlapRatio = MinimumConnectorLabelOverlapRatio
            };
        }
    }
}
