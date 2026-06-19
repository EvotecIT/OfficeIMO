namespace OfficeIMO.Word {
    /// <summary>
    /// Describes the kind of change found while comparing two Word documents.
    /// </summary>
    public enum WordComparisonChangeKind {
        /// <summary>Content exists only in the target document.</summary>
        Inserted,

        /// <summary>Content exists only in the source document.</summary>
        Deleted,

        /// <summary>Content exists in both documents but differs.</summary>
        Modified
    }

    /// <summary>
    /// Describes the document area where a comparison finding was detected.
    /// </summary>
    public enum WordComparisonScope {
        /// <summary>A body paragraph outside a table.</summary>
        Paragraph,

        /// <summary>A table as a whole.</summary>
        Table,

        /// <summary>A table row.</summary>
        TableRow,

        /// <summary>A table cell.</summary>
        TableCell,

        /// <summary>An embedded image payload.</summary>
        Image
    }

    /// <summary>
    /// Represents one machine-readable difference between two Word documents.
    /// </summary>
    public sealed class WordComparisonFinding {
        /// <summary>
        /// Creates a new comparison finding.
        /// </summary>
        /// <param name="scope">Document area where the change was found.</param>
        /// <param name="changeKind">Kind of change found.</param>
        /// <param name="location">Stable, human-readable path such as <c>paragraph[0]</c> or <c>table[1]/row[2]/cell[0]</c>.</param>
        /// <param name="sourceIndex">Index in the source collection when available.</param>
        /// <param name="targetIndex">Index in the target collection when available.</param>
        /// <param name="sourceText">Source text when the finding has textual content.</param>
        /// <param name="targetText">Target text when the finding has textual content.</param>
        /// <param name="message">Short diagnostic message suitable for logs and review reports.</param>
        public WordComparisonFinding(
            WordComparisonScope scope,
            WordComparisonChangeKind changeKind,
            string location,
            int? sourceIndex,
            int? targetIndex,
            string? sourceText,
            string? targetText,
            string message) {
            if (string.IsNullOrWhiteSpace(location)) {
                throw new ArgumentException("Comparison finding location cannot be empty.", nameof(location));
            }

            if (string.IsNullOrWhiteSpace(message)) {
                throw new ArgumentException("Comparison finding message cannot be empty.", nameof(message));
            }

            Scope = scope;
            ChangeKind = changeKind;
            Location = location;
            SourceIndex = sourceIndex;
            TargetIndex = targetIndex;
            SourceText = sourceText;
            TargetText = targetText;
            Message = message;
        }

        /// <summary>Document area where the change was found.</summary>
        public WordComparisonScope Scope { get; }

        /// <summary>Kind of change found.</summary>
        public WordComparisonChangeKind ChangeKind { get; }

        /// <summary>Stable, human-readable path such as <c>paragraph[0]</c> or <c>table[1]/row[2]/cell[0]</c>.</summary>
        public string Location { get; }

        /// <summary>Index in the source collection when available.</summary>
        public int? SourceIndex { get; }

        /// <summary>Index in the target collection when available.</summary>
        public int? TargetIndex { get; }

        /// <summary>Source text when the finding has textual content.</summary>
        public string? SourceText { get; }

        /// <summary>Target text when the finding has textual content.</summary>
        public string? TargetText { get; }

        /// <summary>Short diagnostic message suitable for logs and review reports.</summary>
        public string Message { get; }
    }

    /// <summary>
    /// Machine-readable comparison result produced by <see cref="WordDocumentComparer.CompareStructure(string, string)"/>.
    /// </summary>
    public sealed class WordComparisonResult {
        private readonly List<WordComparisonFinding> _findings = new();

        internal WordComparisonResult(string sourcePath, string targetPath) {
            SourcePath = sourcePath ?? string.Empty;
            TargetPath = targetPath ?? string.Empty;
        }

        /// <summary>Source document path used for the comparison.</summary>
        public string SourcePath { get; }

        /// <summary>Target document path used for the comparison.</summary>
        public string TargetPath { get; }

        /// <summary>All detected findings in deterministic document order.</summary>
        public IReadOnlyList<WordComparisonFinding> Findings => _findings;

        /// <summary>Gets whether any differences were detected.</summary>
        public bool HasChanges => _findings.Count > 0;

        internal void Add(WordComparisonFinding finding) {
            _findings.Add(finding);
        }
    }
}
