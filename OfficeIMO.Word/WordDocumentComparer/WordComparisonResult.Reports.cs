namespace OfficeIMO.Word {
    /// <content>
    /// Compatibility report helpers for <see cref="WordComparisonResult"/>.
    /// </content>
    public sealed partial class WordComparisonResult {
        /// <summary>
        /// Serializes this comparison result to deterministic JSON.
        /// </summary>
        public string ToJson() => WordComparisonReportWriter.ToJson(this);

        /// <summary>
        /// Renders this comparison result as Markdown suitable for review notes and automation logs.
        /// </summary>
        public string ToMarkdown() => WordComparisonReportWriter.ToMarkdown(this);

        /// <summary>
        /// Returns a compact single-line summary suitable for CLI wrappers and CI annotations.
        /// </summary>
        public string ToTextSummary() => WordComparisonReportWriter.ToTextSummary(this);
    }
}
