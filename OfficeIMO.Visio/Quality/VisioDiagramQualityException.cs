using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Exception thrown when a visual quality gate fails.
    /// </summary>
    public sealed class VisioDiagramQualityException : InvalidOperationException {
        /// <summary>
        /// Initializes a new visual quality exception.
        /// </summary>
        public VisioDiagramQualityException(
            IEnumerable<VisioDiagramQualityIssue> issues,
            VisioDiagramQualityIssueSeverity minimumSeverity)
            : base(CreateMessage(issues, minimumSeverity, out IReadOnlyList<VisioDiagramQualityIssue> capturedIssues)) {
            Issues = capturedIssues;
            MinimumSeverity = minimumSeverity;
        }

        /// <summary>
        /// Gets the issues that failed the quality gate.
        /// </summary>
        public IReadOnlyList<VisioDiagramQualityIssue> Issues { get; }

        /// <summary>
        /// Gets the minimum severity used by the quality gate.
        /// </summary>
        public VisioDiagramQualityIssueSeverity MinimumSeverity { get; }

        private static string CreateMessage(
            IEnumerable<VisioDiagramQualityIssue> issues,
            VisioDiagramQualityIssueSeverity minimumSeverity,
            out IReadOnlyList<VisioDiagramQualityIssue> capturedIssues) {
            if (issues == null) throw new ArgumentNullException(nameof(issues));

            capturedIssues = issues.ToList().AsReadOnly();
            string summary = $"Visio diagram quality gate failed with {capturedIssues.Count} issue(s) at or above {minimumSeverity}.";
            if (capturedIssues.Count == 0) {
                return summary;
            }

            return summary + Environment.NewLine + string.Join(Environment.NewLine, capturedIssues.Select(issue => "- " + issue));
        }
    }
}
