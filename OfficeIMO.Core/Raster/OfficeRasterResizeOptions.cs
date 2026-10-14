namespace OfficeIMO.Drawing {
    /// <summary>Requested pixel bounds and sampling for a managed raster resize.</summary>
    /// <remarks>Contain returns fitted pixels without padding. Cover requires both axes and returns a centered crop.
    /// Settings are captured by the resize plan before allocation; later changes do not affect that plan.</remarks>
    public sealed class OfficeRasterResizeOptions {
        /// <summary>Requested width or width bound. At least one axis must be supplied.</summary>
        public int? Width { get; set; }
        /// <summary>Requested height or height bound.</summary>
        public int? Height { get; set; }
        /// <summary>Stretch keeps an omitted source axis; Contain infers it; Cover requires both axes.</summary>
        public OfficeImageFit Fit { get; set; } = OfficeImageFit.Contain;
        /// <summary>Sampling kernel. The editing default is bicubic.</summary>
        public OfficeRasterResamplingMode ResamplingMode { get; set; } = OfficeRasterResamplingMode.Bicubic;
        /// <summary>Color space in which samples are filtered.</summary>
        public OfficeRasterResamplingColorSpace ColorSpace { get; set; } = OfficeRasterResamplingColorSpace.EncodedSrgb;
    }
}
