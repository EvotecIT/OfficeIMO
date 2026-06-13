namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Controls whether Word visual content that can be rendered by OfficeIMO is emitted as Markdown images.
    /// </summary>
    public enum MarkdownVisualFallbackMode {
        /// <summary>Do not render visual fallbacks. Use unsupported-content handling instead.</summary>
        None,

        /// <summary>Render supported visual content as SVG data URI images.</summary>
        SvgDataUri,

        /// <summary>Render supported visual content as SVG files and reference them from Markdown image links.</summary>
        SvgFile
    }
}
