namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Controls how markdown image dimensions are normalized before insertion into Word.
    /// </summary>
    public sealed class MarkdownImageLayoutOptions {
        /// <summary>
        /// Selects how container width should constrain markdown images.
        /// </summary>
        public MarkdownImageFitMode FitMode { get; set; } = MarkdownImageFitMode.None;

        /// <summary>
        /// Defines whether markdown width/height hints are applied before or after layout constraints.
        /// </summary>
        public MarkdownImageHintPrecedence HintPrecedence { get; set; } = MarkdownImageHintPrecedence.MarkdownThenLayout;

        /// <summary>
        /// Optional hard cap (pixels) for image width.
        /// </summary>
        public double? MaxWidthPixels { get; set; }

        /// <summary>
        /// Optional hard cap (pixels) for image height.
        /// </summary>
        public double? MaxHeightPixels { get; set; }

        /// <summary>
        /// Optional hard cap expressed as percent of available content width (for example 100, 85, 50).
        /// When set, this is evaluated against context width (or page content width when context is unavailable).
        /// </summary>
        public double? MaxWidthPercentOfContent { get; set; }

        /// <summary>
        /// When <c>false</c>, explicit markdown hints larger than natural dimensions are clamped down.
        /// </summary>
        public bool AllowUpscale { get; set; }

        /// <summary>
        /// When enabled, SVG sources are rasterized to PNG before insertion to improve downstream compatibility
        /// (for example external viewers and conversion pipelines that do not fully support inline SVG).
        /// </summary>
        public bool PreferRasterizeSvgForWord { get; set; }

        /// <summary>
        /// Target rasterization DPI used when <see cref="PreferRasterizeSvgForWord"/> is enabled.
        /// Default is 144 DPI for better readability than 96 DPI while keeping file sizes moderate.
        /// </summary>
        public int SvgRasterizationDpi { get; set; } = 144;
    }

    /// <summary>
    /// Selects image fitting behavior for markdown conversion.
    /// </summary>
    public enum MarkdownImageFitMode {
        /// <summary>
        /// Do not apply container-based width fitting.
        /// </summary>
        None = 0,
        /// <summary>
        /// Fit to section content width (page width minus margins).
        /// </summary>
        PageContentWidth = 1,
        /// <summary>
        /// Fit to section content width minus structural indentation (quotes/lists).
        /// </summary>
        ContextContentWidth = 2
    }

    /// <summary>
    /// Controls ordering between markdown size hints and layout constraints.
    /// </summary>
    public enum MarkdownImageHintPrecedence {
        /// <summary>
        /// Apply markdown width/height hints first, then clamp using layout constraints.
        /// </summary>
        MarkdownThenLayout = 0,
        /// <summary>
        /// Start from natural dimensions, then apply markdown hints and final layout clamping.
        /// </summary>
        LayoutThenMarkdown = 1
    }

    /// <summary>
    /// Provides per-image layout diagnostics emitted during markdown conversion.
    /// </summary>
    public sealed class MarkdownImageLayoutDiagnostic {
        /// <summary>
        /// Image source path or URL.
        /// </summary>
        public string Source { get; set; } = string.Empty;

        /// <summary>
        /// Logical conversion context (for example block-local, block-remote).
        /// </summary>
        public string Context { get; set; } = string.Empty;

        /// <summary>
        /// Requested width hint from markdown or caller.
        /// </summary>
        public double? RequestedWidthPixels { get; set; }

        /// <summary>
        /// Requested height hint from markdown or caller.
        /// </summary>
        public double? RequestedHeightPixels { get; set; }

        /// <summary>
        /// Natural image width when available.
        /// </summary>
        public double? NaturalWidthPixels { get; set; }

        /// <summary>
        /// Natural image height when available.
        /// </summary>
        public double? NaturalHeightPixels { get; set; }

        /// <summary>
        /// Effective width limit applied by converter logic.
        /// </summary>
        public double? EffectiveMaxWidthPixels { get; set; }

        /// <summary>
        /// Effective height limit applied by converter logic.
        /// </summary>
        public double? EffectiveMaxHeightPixels { get; set; }

        /// <summary>
        /// Final width selected for insertion.
        /// </summary>
        public double? FinalWidthPixels { get; set; }

        /// <summary>
        /// Final height selected for insertion.
        /// </summary>
        public double? FinalHeightPixels { get; set; }

        /// <summary>
        /// Indicates that at least one layout constraint changed the dimensions.
        /// </summary>
        public bool ScaledByLayout { get; set; }

        /// <summary>
        /// Indicates that the source image was SVG and was rasterized before insertion.
        /// </summary>
        public bool RasterizedFromSvg { get; set; }
    }
}
