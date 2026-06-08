namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Options for rendering Markdown into an existing Word template document.
    /// </summary>
    public class MarkdownToWordTemplateOptions : MarkdownToWordOptions {
        /// <summary>
        /// Creates template insertion options. Front matter is hidden by default because
        /// template workflows usually consume it as metadata rather than body content.
        /// </summary>
        public MarkdownToWordTemplateOptions() {
            RenderFrontMatter = false;
        }

        /// <summary>
        /// Tag of a block content control that marks the Markdown insertion point.
        /// </summary>
        public string? ContentControlTag { get; set; }

        /// <summary>
        /// Alias of a block content control that marks the Markdown insertion point.
        /// </summary>
        public string? ContentControlAlias { get; set; }

        /// <summary>
        /// Bookmark name that marks the Markdown insertion point.
        /// </summary>
        public string? BookmarkName { get; set; }

        /// <summary>
        /// Replaces the target placeholder element after inserting Markdown.
        /// </summary>
        public bool ReplacePlaceholder { get; set; } = true;
    }
}
