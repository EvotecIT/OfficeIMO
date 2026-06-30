namespace OfficeIMO.Word {
    /// <summary>
    /// Selects the DOCX shape produced by <see cref="WordDocumentComparer.CreateRedlineDocument(string, string, string, WordComparisonRedlineOptions?)"/>.
    /// </summary>
    public enum WordComparisonRedlineMode {
        /// <summary>
        /// Creates a standalone review artifact with a summary, findings table, and tracked snippets.
        /// </summary>
        ReportArtifact,

        /// <summary>
        /// Clones the target document and applies supported tracked changes in place.
        /// </summary>
        InPlaceTarget
    }

    /// <summary>
    /// Controls redline document generation for structured Word document comparisons.
    /// </summary>
    public sealed class WordComparisonRedlineOptions {
        /// <summary>
        /// Gets or sets the generated redline document shape.
        /// </summary>
        public WordComparisonRedlineMode Mode { get; set; } = WordComparisonRedlineMode.ReportArtifact;

        /// <summary>
        /// Gets or sets the author used for generated tracked insertions and deletions.
        /// </summary>
        public string Author { get; set; } = "OfficeIMO";

        /// <summary>
        /// Gets or sets the revision timestamp used for generated tracked insertions and deletions.
        /// </summary>
        public DateTime? DateTime { get; set; }

        /// <summary>
        /// Gets or sets the structured comparison options used before redline generation.
        /// </summary>
        public WordComparisonOptions? ComparisonOptions { get; set; }

        /// <summary>
        /// Gets or sets whether the generated document includes a summary section.
        /// </summary>
        public bool IncludeSummary { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the generated document includes a findings table.
        /// </summary>
        public bool IncludeFindingsTable { get; set; } = true;

        /// <summary>
        /// Gets or sets whether text-bearing findings are emitted as tracked insertions and deletions.
        /// </summary>
        public bool TrackTextFindings { get; set; } = true;

        /// <summary>
        /// Gets or sets whether feature/metadata findings such as fields, content controls, bookmarks, hyperlinks, lists, and images are emitted as tracked revisions.
        /// </summary>
        public bool TrackFeatureFindings { get; set; } = true;

        /// <summary>
        /// Gets or sets whether comment and revision comparison findings are emitted as tracked revisions.
        /// </summary>
        public bool TrackReviewFindings { get; set; } = true;

        /// <summary>
        /// Gets or sets whether formatting-only findings are emitted as tracked revisions.
        /// </summary>
        public bool TrackFormattingFindings { get; set; } = true;
    }
}
