using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Exception thrown when a Visio showcase proof bundle fails artifact validation.
    /// </summary>
    public sealed class VisioShowcaseValidationException : InvalidOperationException {
        /// <summary>
        /// Initializes a new showcase validation exception.
        /// </summary>
        public VisioShowcaseValidationException(IEnumerable<VisioShowcaseValidationIssue> issues)
            : base(CreateMessage(issues, out IReadOnlyList<VisioShowcaseValidationIssue> capturedIssues)) {
            Issues = capturedIssues;
        }

        /// <summary>Issues that failed the showcase validation gate.</summary>
        public IReadOnlyList<VisioShowcaseValidationIssue> Issues { get; }

        private static string CreateMessage(
            IEnumerable<VisioShowcaseValidationIssue> issues,
            out IReadOnlyList<VisioShowcaseValidationIssue> capturedIssues) {
            if (issues == null) throw new ArgumentNullException(nameof(issues));

            capturedIssues = issues.ToList().AsReadOnly();
            string summary = $"Visio showcase artifact validation failed with {capturedIssues.Count} issue(s).";
            if (capturedIssues.Count == 0) {
                return summary;
            }

            return summary + Environment.NewLine + string.Join(Environment.NewLine, capturedIssues.Select(issue => "- " + issue));
        }
    }
}
