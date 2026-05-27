namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for generating OfficeIMO.Visio gallery documents.
    /// </summary>
    public sealed class VisioGalleryOptions {
        /// <summary>Whether generated packages should be structurally validated.</summary>
        public bool ValidatePackage { get; set; } = true;

        /// <summary>Whether generated pages should be analyzed for visual quality issues.</summary>
        public bool AnalyzeVisualQuality { get; set; } = true;

        /// <summary>Whether generated documents should also be opened with Microsoft Visio desktop when available.</summary>
        public bool ValidateWithVisioDesktop { get; set; }

        /// <summary>Whether missing Microsoft Visio desktop automation should make gallery results unclean.</summary>
        public bool RequireVisioDesktop { get; set; }

        /// <summary>Optional desktop validation steps, such as round-trip save-copy and SVG/PNG/PDF export proof files.</summary>
        public VisioDesktopValidationOptions? DesktopValidationOptions { get; set; }

        /// <summary>Visual quality options used when <see cref="AnalyzeVisualQuality"/> is enabled.</summary>
        public VisioDiagramQualityOptions QualityOptions { get; set; } = new VisioDiagramQualityOptions {
            CheckShapeOverlaps = true,
            CheckConnectorShapeIntersections = true,
            CheckConnectorLabels = true
        };
    }
}
