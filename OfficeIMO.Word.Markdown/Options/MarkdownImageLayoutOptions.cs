namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Controls how markdown image dimensions are normalized before insertion into Word.
    /// </summary>
    public sealed class WordMarkdownImageLayoutOptions {
        internal WordMarkdownImageLayoutOptions CloneForConversion() => (WordMarkdownImageLayoutOptions)MemberwiseClone();

        /// <summary>
        /// Selects how container width should constrain markdown images.
        /// </summary>
        public WordMarkdownImageFitMode FitMode { get; set; } = WordMarkdownImageFitMode.None;

        /// <summary>
        /// Defines whether markdown width/height hints are applied before or after layout constraints.
        /// </summary>
        public WordMarkdownImageHintPrecedence HintPrecedence { get; set; } = WordMarkdownImageHintPrecedence.MarkdownThenLayout;

        /// <summary>
        /// Optional physical hard cap for image width.
        /// </summary>
        public global::OfficeIMO.OfficeLength? MaxWidth { get; set; }

        /// <summary>
        /// Optional physical hard cap for image height.
        /// </summary>
        public global::OfficeIMO.OfficeLength? MaxHeight { get; set; }

        /// <summary>
        /// Optional hard cap expressed as percent of available content width (for example 100, 85, 50).
        /// When set, this is evaluated against context width (or page content width when context is unavailable).
        /// </summary>
        public double? MaxWidthPercentOfContent { get; set; }

        /// <summary>
        /// When <c>false</c>, explicit markdown hints larger than natural dimensions are clamped down.
        /// </summary>
        public bool AllowUpscale { get; set; }
    }

    /// <summary>
    /// Selects image fitting behavior for markdown conversion.
    /// </summary>
    public enum WordMarkdownImageFitMode {
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
    public enum WordMarkdownImageHintPrecedence {
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
    public sealed class WordMarkdownImageLayoutDiagnostic {
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
        public global::OfficeIMO.OfficeLength? RequestedWidth { get; set; }

        /// <summary>
        /// Requested height hint from markdown or caller.
        /// </summary>
        public global::OfficeIMO.OfficeLength? RequestedHeight { get; set; }

        /// <summary>
        /// Natural image width interpreted at 96 DPI when available.
        /// </summary>
        public global::OfficeIMO.OfficeLength? NaturalWidth { get; set; }

        /// <summary>
        /// Natural image height interpreted at 96 DPI when available.
        /// </summary>
        public global::OfficeIMO.OfficeLength? NaturalHeight { get; set; }

        /// <summary>
        /// Effective width limit applied by converter logic.
        /// </summary>
        public global::OfficeIMO.OfficeLength? EffectiveMaxWidth { get; set; }

        /// <summary>
        /// Effective height limit applied by converter logic.
        /// </summary>
        public global::OfficeIMO.OfficeLength? EffectiveMaxHeight { get; set; }

        /// <summary>
        /// Final width selected for insertion.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FinalWidth { get; set; }

        /// <summary>
        /// Final height selected for insertion.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FinalHeight { get; set; }

        /// <summary>
        /// Indicates that at least one layout constraint changed the dimensions.
        /// </summary>
        public bool ScaledByLayout { get; set; }
    }
}
