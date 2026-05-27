using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Summarizes dependency-free visual quality analysis for a page or document.
    /// </summary>
    public sealed class VisioDiagramQualityReport {
        /// <summary>
        /// Initializes a new visual quality report.
        /// </summary>
        public VisioDiagramQualityReport(IEnumerable<VisioDiagramQualityIssue> issues) {
            if (issues == null) throw new ArgumentNullException(nameof(issues));

            Issues = issues.ToList().AsReadOnly();
            InformationCount = Issues.Count(issue => issue.Severity == VisioDiagramQualityIssueSeverity.Information);
            WarningCount = Issues.Count(issue => issue.Severity == VisioDiagramQualityIssueSeverity.Warning);
            ErrorCount = Issues.Count(issue => issue.Severity == VisioDiagramQualityIssueSeverity.Error);
        }

        /// <summary>
        /// Gets all quality issues.
        /// </summary>
        public IReadOnlyList<VisioDiagramQualityIssue> Issues { get; }

        /// <summary>
        /// Gets the number of informational issues.
        /// </summary>
        public int InformationCount { get; }

        /// <summary>
        /// Gets the number of warnings.
        /// </summary>
        public int WarningCount { get; }

        /// <summary>
        /// Gets the number of errors.
        /// </summary>
        public int ErrorCount { get; }

        /// <summary>
        /// Gets whether the report has no warnings or errors.
        /// </summary>
        public bool IsClean => !HasIssuesAtOrAbove(VisioDiagramQualityIssueSeverity.Warning);

        /// <summary>
        /// Gets issues at or above the given severity.
        /// </summary>
        public IReadOnlyList<VisioDiagramQualityIssue> GetIssuesAtOrAbove(VisioDiagramQualityIssueSeverity minimumSeverity) {
            return Issues
                .Where(issue => issue.Severity >= minimumSeverity)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
        /// Gets whether the report has issues at or above the given severity.
        /// </summary>
        public bool HasIssuesAtOrAbove(VisioDiagramQualityIssueSeverity minimumSeverity) {
            return Issues.Any(issue => issue.Severity >= minimumSeverity);
        }

        /// <summary>
        /// Throws when this report contains issues at or above the given severity.
        /// </summary>
        public void EnsureClean(VisioDiagramQualityIssueSeverity minimumSeverity = VisioDiagramQualityIssueSeverity.Warning) {
            IReadOnlyList<VisioDiagramQualityIssue> blockingIssues = GetIssuesAtOrAbove(minimumSeverity);
            if (blockingIssues.Count > 0) {
                throw new VisioDiagramQualityException(blockingIssues, minimumSeverity);
            }
        }

        /// <inheritdoc />
        public override string ToString() {
            if (Issues.Count == 0) {
                return "No visual quality issues.";
            }

            StringBuilder builder = new();
            builder.Append("Visual quality issues: ");
            builder.Append(ErrorCount);
            builder.Append(" error(s), ");
            builder.Append(WarningCount);
            builder.Append(" warning(s), ");
            builder.Append(InformationCount);
            builder.Append(" information item(s).");

            foreach (VisioDiagramQualityIssue issue in Issues) {
                builder.AppendLine();
                builder.Append("- ");
                builder.Append(issue.ToString());
            }

            return builder.ToString();
        }
    }
}
