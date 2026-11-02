using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>Controls worksheet image geometry, source ownership and native drawing metadata.</summary>
    public sealed class ExcelImageOptions {
        /// <summary>Gets or sets the exact width. Omitted byte-image widths are 96 pixels; loaded sources use detected dimensions.</summary>
        public OfficeLength? Width { get; set; }
        /// <summary>Gets or sets the exact height. Omitted byte-image heights are 32 pixels; loaded sources use detected dimensions.</summary>
        public OfficeLength? Height { get; set; }
        /// <summary>Gets or sets a positive percentage of detected source dimensions. Cannot be combined with explicit dimensions or a range anchor.</summary>
        public double? ScalePercent { get; set; }
        /// <summary>Gets or sets the exact horizontal offset of the starting cell marker, including signed native offsets.</summary>
        public OfficeLength OffsetX { get; set; }
        /// <summary>Gets or sets the exact vertical offset of the starting cell marker, including signed native offsets.</summary>
        public OfficeLength OffsetY { get; set; }
        /// <summary>Gets or sets the exact horizontal offset of a two-cell ending marker.</summary>
        public OfficeLength EndOffsetX { get; set; }
        /// <summary>Gets or sets the exact vertical offset of a two-cell ending marker.</summary>
        public OfficeLength EndOffsetY { get; set; }
        /// <summary>Gets or sets how a two-cell anchor responds to cell moves and resizes.</summary>
        public ExcelImagePlacement Placement { get; set; } = ExcelImagePlacement.MoveAndSize;
        /// <summary>Gets or sets the drawing name used by Excel's selection pane.</summary>
        public string? Name { get; set; }
        /// <summary>Gets or sets the alternative text description.</summary>
        public string? Description { get; set; }
        /// <summary>Gets or sets the alternative text title.</summary>
        public string? Title { get; set; }
        /// <summary>Gets or sets whether the picture's aspect ratio is locked in Excel. Does not infer an omitted dimension.</summary>
        public bool LockAspectRatio { get; set; } = true;
        /// <summary>Gets or sets the clockwise rotation in degrees.</summary>
        public double RotationDegrees { get; set; }
        /// <summary>Gets or sets the MIME type for native byte payloads. Loaded sources use their validated content type.</summary>
        public string ContentType { get; set; } = "image/png";
        /// <summary>Gets or sets whether a caller-owned stream stays open on success, cancellation and failure. Its seekable position is preserved.</summary>
        public bool LeaveOpen { get; set; } = true;
        /// <summary>Gets or sets whether identical bytes of the same native type reuse an existing image part in this worksheet drawing. Defaults to separate parts.</summary>
        public bool Deduplicate { get; set; }
        /// <summary>Gets or sets the existing network policy for HTTP or HTTPS image loading.</summary>
        public OfficeRemoteImageLoadOptions? RemoteImageOptions { get; set; }

        internal ExcelImageOptions Snapshot() => (ExcelImageOptions)MemberwiseClone();
    }
}
