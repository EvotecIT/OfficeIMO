namespace OfficeIMO.Visio {
    /// <summary>
    /// Describes a problem found while validating a generated Visio showcase proof bundle.
    /// </summary>
    public sealed class VisioShowcaseValidationIssue {
        /// <summary>
        /// Initializes a new showcase validation issue.
        /// </summary>
        public VisioShowcaseValidationIssue(
            string kind,
            string message,
            string? relativePath = null,
            string? diagramName = null) {
            Kind = kind;
            Message = message;
            RelativePath = relativePath;
            DiagramName = diagramName;
        }

        /// <summary>Stable issue kind.</summary>
        public string Kind { get; }

        /// <summary>Human-readable validation message.</summary>
        public string Message { get; }

        /// <summary>Showcase-relative artifact path, when the issue is tied to a file.</summary>
        public string? RelativePath { get; }

        /// <summary>Showcase diagram name, when the issue is tied to a diagram group.</summary>
        public string? DiagramName { get; }

        /// <inheritdoc />
        public override string ToString() {
            string path = string.IsNullOrWhiteSpace(RelativePath) ? string.Empty : $" [{RelativePath}]";
            string diagram = string.IsNullOrWhiteSpace(DiagramName) ? string.Empty : $" on '{DiagramName}'";
            return $"{Kind}{diagram}{path}: {Message}";
        }
    }
}
