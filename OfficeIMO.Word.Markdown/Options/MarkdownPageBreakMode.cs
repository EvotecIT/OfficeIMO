namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Controls how Word page breaks are represented when exporting to Markdown.
    /// </summary>
    public enum MarkdownPageBreakMode {
        /// <summary>
        /// Emit an OfficeIMO semantic fenced block that can be parsed back into a Word page break.
        /// </summary>
        SemanticBlock = 0,

        /// <summary>
        /// Emit an HTML block using a CSS page-break marker.
        /// </summary>
        Html = 1,

        /// <summary>
        /// Emit a Markdown thematic break (<c>---</c>).
        /// </summary>
        HorizontalRule = 2,

        /// <summary>
        /// Do not emit page-break markers.
        /// </summary>
        Omit = 3
    }
}
