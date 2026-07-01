namespace OfficeIMO.Word {
    /// <summary>
    /// Identifies the operation applied to tracked revisions.
    /// </summary>
    public enum WordRevisionOperationKind {
        /// <summary>Accept matching revisions.</summary>
        Accept,
        /// <summary>Reject matching revisions.</summary>
        Reject
    }

    /// <summary>
    /// Filters tracked revisions for scoped accept/reject operations.
    /// </summary>
    public sealed class WordRevisionFilter {
        /// <summary>Matches revisions authored by this person using case-insensitive comparison.</summary>
        public string? Author { get; set; }

        /// <summary>Matches a specific revision id.</summary>
        public string? RevisionId { get; set; }

        /// <summary>Matches a specific revision type.</summary>
        public WordReviewRevisionType? RevisionType { get; set; }

        /// <summary>Matches revisions at or after this timestamp.</summary>
        public DateTime? DateFrom { get; set; }

        /// <summary>Matches revisions at or before this timestamp.</summary>
        public DateTime? DateTo { get; set; }

        /// <summary>Matches revisions in a specific document part kind.</summary>
        public WordReviewLocationKind? LocationKind { get; set; }

        /// <summary>Matches revisions in a specific package part URI.</summary>
        public string? PartUri { get; set; }

        /// <summary>When set, matches whether the revision is inside a table.</summary>
        public bool? IsInTable { get; set; }

        /// <summary>When set, matches whether the revision is inside a content control.</summary>
        public bool? IsInContentControl { get; set; }

        /// <summary>When set, matches whether the revision is inside a text box.</summary>
        public bool? IsInTextBox { get; set; }

        /// <summary>Gets a filter that matches all revisions.</summary>
        public static WordRevisionFilter All() {
            return new WordRevisionFilter();
        }
    }

    /// <summary>
    /// Result from a scoped tracked-revision operation.
    /// </summary>
    public sealed class WordRevisionOperationReport {
        internal WordRevisionOperationReport(WordRevisionOperationKind operation, IReadOnlyList<WordRevisionInfo> matchedRevisions) {
            Operation = operation;
            MatchedRevisions = matchedRevisions.ToArray();
        }

        /// <summary>Gets the operation that was applied.</summary>
        public WordRevisionOperationKind Operation { get; }

        /// <summary>Gets revisions that matched before the document was mutated.</summary>
        public IReadOnlyList<WordRevisionInfo> MatchedRevisions { get; }

        /// <summary>Gets the number of matching revisions processed.</summary>
        public int MatchedCount => MatchedRevisions.Count;
    }
}
