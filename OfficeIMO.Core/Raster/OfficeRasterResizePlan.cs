namespace OfficeIMO.Drawing {
    /// <summary>Immutable pixel geometry and peak managed-storage plan for one raster resize.</summary>
    /// <remarks>Contain dimensions round to the nearest pixel with midpoint ties away from zero and remain within supplied bounds.
    /// Cover dimensions round upward; its centered crop puts an odd extra pixel on the right or bottom.
    /// The plan includes the independent-copy path, sampling scratch and tables, and any resized intermediate and crop output.</remarks>
    public sealed class OfficeRasterResizePlan {
        internal OfficeRasterResizePlan(int sourceWidth, int sourceHeight, int resizeWidth, int resizeHeight,
            int width, int height, OfficeImageFit fit, OfficeRasterResamplingMode mode,
            OfficeRasterResamplingColorSpace colorSpace, long workingSetBytes) {
            SourceWidth = sourceWidth; SourceHeight = sourceHeight;
            ResizeWidth = resizeWidth; ResizeHeight = resizeHeight;
            Width = width; Height = height;
            CropX = (resizeWidth - width) / 2; CropY = (resizeHeight - height) / 2;
            Fit = fit; ResamplingMode = mode; ColorSpace = colorSpace;
            WorkingSetBytes = workingSetBytes;
            AdditionalWorkingBytes = workingSetBytes - ((long)sourceWidth * sourceHeight + (long)width * height) * 4L;
        }
        /// <summary>Required source width.</summary>
        public int SourceWidth { get; }
        /// <summary>Required source height.</summary>
        public int SourceHeight { get; }
        /// <summary>Width of the sampled intermediate before an optional crop.</summary>
        public int ResizeWidth { get; }
        /// <summary>Height of the sampled intermediate before an optional crop.</summary>
        public int ResizeHeight { get; }
        /// <summary>Final output width.</summary>
        public int Width { get; }
        /// <summary>Final output height.</summary>
        public int Height { get; }
        /// <summary>Left pixel of the final crop in the sampled intermediate.</summary>
        public int CropX { get; }
        /// <summary>Top pixel of the final crop in the sampled intermediate.</summary>
        public int CropY { get; }
        /// <summary>Captured fit behavior.</summary>
        public OfficeImageFit Fit { get; }
        /// <summary>Captured sampling kernel.</summary>
        public OfficeRasterResamplingMode ResamplingMode { get; }
        /// <summary>Captured filtering color space.</summary>
        public OfficeRasterResamplingColorSpace ColorSpace { get; }
        /// <summary>Peak source, output, temporary, table and fixed-overhead storage in bytes.</summary>
        public long WorkingSetBytes { get; }
        /// <summary>Peak storage beyond source and final-output RGBA buffers; frame operations retain this allowance alongside all frames.</summary>
        public long AdditionalWorkingBytes { get; }
    }
}
