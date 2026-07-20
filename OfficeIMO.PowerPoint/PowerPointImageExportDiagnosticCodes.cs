namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Stable diagnostic codes emitted by PowerPoint image export.
    /// </summary>
    public static class PowerPointImageExportDiagnosticCodes {
        /// <summary>A PowerPoint shape or one of its visual features could not be projected exactly.</summary>
        public const string UnsupportedShape = "unsupported-powerpoint-shape";

        /// <summary>A PowerPoint image was omitted because its projected bounds exceed the slide canvas.</summary>
        public const string ImageBoundsUnsupported = "unsupported-powerpoint-image-bounds";

        /// <summary>The slide background color could not be parsed and the configured fallback was used.</summary>
        public const string InvalidSlideBackgroundColor = "invalid-slide-background-color";

        /// <summary>The slide background gradient could not be parsed and the configured fallback was used.</summary>
        public const string InvalidSlideBackgroundGradient = "invalid-slide-background-gradient";

        /// <summary>The slide background kind is unsupported and the configured fallback was used.</summary>
        public const string UnsupportedSlideBackground = "unsupported-slide-background";

        /// <summary>The slide background image could not be read and the configured fallback was used.</summary>
        public const string InvalidSlideBackgroundImage = "invalid-slide-background-image";
    }
}
