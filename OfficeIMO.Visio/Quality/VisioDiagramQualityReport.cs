using System;
using System.Collections.Generic;
using System.Globalization;
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

        /// <summary>
        /// Renders a deterministic key-value proof artifact suitable for CI and showcase review bundles.
        /// </summary>
        public string ToText() {
            StringBuilder builder = new();
            builder.Append("quality.isClean=");
            builder.AppendLine(IsClean ? "true" : "false");
            builder.Append("quality.issueCount=");
            builder.AppendLine(Issues.Count.ToString(CultureInfo.InvariantCulture));
            builder.Append("quality.errorCount=");
            builder.AppendLine(ErrorCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("quality.warningCount=");
            builder.AppendLine(WarningCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("quality.informationCount=");
            builder.AppendLine(InformationCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("quality.issueKinds=");
            builder.AppendLine(string.Join(",", Issues.Select(issue => issue.Kind).Distinct(StringComparer.Ordinal).OrderBy(kind => kind, StringComparer.Ordinal)));

            int index = 0;
            foreach (VisioDiagramQualityIssue issue in Issues.OrderBy(issue => issue.PageName, StringComparer.Ordinal)
                         .ThenBy(issue => issue.Kind, StringComparer.Ordinal)
                         .ThenBy(issue => issue.ShapeId, StringComparer.Ordinal)
                         .ThenBy(issue => issue.OtherShapeId, StringComparer.Ordinal)
                         .ThenBy(issue => issue.ConnectorId, StringComparer.Ordinal)
                         .ThenBy(issue => issue.OtherConnectorId, StringComparer.Ordinal)
                         .ThenBy(issue => issue.Message, StringComparer.Ordinal)) {
                string prefix = "quality.issue[" + index.ToString(CultureInfo.InvariantCulture) + "].";
                builder.Append(prefix);
                builder.Append("severity=");
                builder.AppendLine(issue.Severity.ToString());
                builder.Append(prefix);
                builder.Append("kind=");
                builder.AppendLine(issue.Kind);
                AppendOptionalProofValue(builder, prefix, "page", issue.PageName);
                AppendOptionalProofValue(builder, prefix, "shape", issue.ShapeId);
                AppendOptionalProofValue(builder, prefix, "otherShape", issue.OtherShapeId);
                AppendOptionalProofValue(builder, prefix, "connector", issue.ConnectorId);
                AppendOptionalProofValue(builder, prefix, "otherConnector", issue.OtherConnectorId);
                builder.Append(prefix);
                builder.Append("message=");
                builder.AppendLine(NormalizeProofValue(issue.Message));
                index++;
            }

            return builder.ToString();
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

        private static void AppendOptionalProofValue(StringBuilder builder, string prefix, string name, string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return;
            }

            builder.Append(prefix);
            builder.Append(name);
            builder.Append('=');
            builder.AppendLine(NormalizeProofValue(value!));
        }

        private static string NormalizeProofValue(string value) {
            return value
                .Replace("\r\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\n", "\\n");
        }
    }
}
