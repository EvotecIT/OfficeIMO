namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Controls how Word content without a native Markdown representation is exported.
    /// </summary>
    public enum MarkdownUnsupportedContentMode {
        /// <summary>
        /// Emit a non-fatal warning through <see cref="WordToMarkdownOptions.OnWarning"/> and omit the content.
        /// </summary>
        WarnOnly = 0,

        /// <summary>
        /// Emit a visible Markdown placeholder paragraph and a non-fatal warning.
        /// </summary>
        Placeholder = 1,

        /// <summary>
        /// Emit an HTML comment marker and a non-fatal warning.
        /// </summary>
        HtmlComment = 2,

        /// <summary>
        /// Omit the content without warning.
        /// </summary>
        Omit = 3
    }
}
