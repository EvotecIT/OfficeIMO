using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for high-level diagram cleanup before saving or validating a Visio document.
    /// </summary>
    public sealed class VisioDiagramPolishOptions {
        /// <summary>Whether top-level shapes should be resized to fit their text.</summary>
        public bool ResizeShapesToText { get; set; }

        /// <summary>Whether connector label boxes should be resized to fit their text.</summary>
        public bool ResizeConnectorLabelsToText { get; set; } = true;

        /// <summary>Whether connector labels should be moved away from page edges, unrelated shapes, and other labels.</summary>
        public bool ResolveConnectorLabelOverlaps { get; set; } = true;

        /// <summary>Whether top-level shapes should be moved apart when their bounds overlap.</summary>
        public bool ResolveShapeOverlaps { get; set; }

        /// <summary>Whether deterministic connector routes should be nudged around unrelated shapes.</summary>
        public bool ResolveConnectorShapeIntersections { get; set; }

        /// <summary>Whether the page should be moved and optionally resized around the polished content.</summary>
        public bool FitToContent { get; set; } = true;

        /// <summary>Whether <see cref="VisioDocument"/> polish should include background pages.</summary>
        public bool IncludeBackgroundPages { get; set; }

        /// <summary>Whether page fitting should resize the page.</summary>
        public bool ResizePage { get; set; } = true;

        /// <summary>Horizontal content margin used by page fitting, in inches.</summary>
        public double FitHorizontalMargin { get; set; } = 0.6D;

        /// <summary>Vertical content margin used by page fitting, in inches.</summary>
        public double FitVerticalMargin { get; set; } = 0.45D;

        /// <summary>Optional font used when resizing shape text boxes.</summary>
        public OfficeFontInfo? ShapeFontInfo { get; set; }

        /// <summary>Horizontal padding used when resizing shape text boxes, in inches.</summary>
        public double ShapeHorizontalPadding { get; set; } = 0.25D;

        /// <summary>Vertical padding used when resizing shape text boxes, in inches.</summary>
        public double ShapeVerticalPadding { get; set; } = 0.14D;

        /// <summary>Minimum shape width used when resizing shape text boxes, in inches.</summary>
        public double MinimumShapeWidth { get; set; } = 0.5D;

        /// <summary>Minimum shape height used when resizing shape text boxes, in inches.</summary>
        public double MinimumShapeHeight { get; set; } = 0.3D;

        /// <summary>Optional font used when resizing connector label boxes.</summary>
        public OfficeFontInfo? ConnectorLabelFontInfo { get; set; }

        /// <summary>Horizontal padding used when resizing connector label boxes, in inches.</summary>
        public double ConnectorLabelHorizontalPadding { get; set; } = 0.12D;

        /// <summary>Vertical padding used when resizing connector label boxes, in inches.</summary>
        public double ConnectorLabelVerticalPadding { get; set; } = 0.06D;

        /// <summary>Minimum connector label width, in inches.</summary>
        public double MinimumConnectorLabelWidth { get; set; } = 0.45D;

        /// <summary>Minimum connector label height, in inches.</summary>
        public double MinimumConnectorLabelHeight { get; set; } = 0.22D;

        /// <summary>Maximum connector label width used for word wrapping, in inches.</summary>
        public double? MaximumConnectorLabelWidth { get; set; } = 1.6D;

        /// <summary>Search step used when moving connector labels, in inches.</summary>
        public double ConnectorLabelStep { get; set; } = 0.18D;

        /// <summary>Number of search rings to try when moving connector labels.</summary>
        public int ConnectorLabelMaxAttempts { get; set; } = 12;

        /// <summary>Padding added around obstacle shapes when rerouting connectors, in inches.</summary>
        public double ConnectorRoutingObstaclePadding { get; set; } = 0.15D;

        /// <summary>Number of positive and negative connector routing lanes to try on each axis.</summary>
        public int ConnectorRoutingMaxLanes { get; set; } = 12;

        /// <summary>Search step used when moving overlapping shapes, in inches.</summary>
        public double ShapeOverlapStep { get; set; } = 0.25D;

        /// <summary>Number of search rings to try when moving overlapping shapes.</summary>
        public int ShapeOverlapMaxAttempts { get; set; } = 24;

        /// <summary>Whether container and background surface shapes should participate in shape overlap cleanup.</summary>
        public bool IncludeContainersInShapeOverlapResolution { get; set; }

        /// <summary>Whether connector labels should avoid unrelated shapes.</summary>
        public bool AvoidConnectorLabelShapeOverlaps { get; set; } = true;

        /// <summary>Whether connector labels should avoid labels that have already been placed.</summary>
        public bool AvoidConnectorLabelOverlaps { get; set; } = true;
    }
}
