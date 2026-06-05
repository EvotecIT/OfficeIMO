using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Summarizes validation of generated Visio showcase packages, previews, and grouped diagram metadata.
    /// </summary>
    public sealed class VisioShowcaseValidationReport {
        /// <summary>
        /// Initializes a new showcase validation report.
        /// </summary>
        public VisioShowcaseValidationReport(IEnumerable<VisioShowcaseValidationIssue> issues) {
            if (issues == null) throw new ArgumentNullException(nameof(issues));
            Issues = issues.ToList().AsReadOnly();
        }

        /// <summary>Validation issues found in the showcase proof bundle.</summary>
        public IReadOnlyList<VisioShowcaseValidationIssue> Issues { get; }

        /// <summary>Whether no validation issues were found.</summary>
        public bool IsClean => Issues.Count == 0;

        /// <summary>
        /// Throws when the showcase proof bundle has validation issues.
        /// </summary>
        public void EnsureClean() {
            if (!IsClean) {
                throw new VisioShowcaseValidationException(Issues);
            }
        }

        /// <inheritdoc />
        public override string ToString() {
            if (Issues.Count == 0) {
                return "No Visio showcase artifact validation issues.";
            }

            StringBuilder builder = new();
            builder.Append("Visio showcase artifact validation found ");
            builder.Append(Issues.Count);
            builder.Append(" issue(s).");
            foreach (VisioShowcaseValidationIssue issue in Issues) {
                builder.AppendLine();
                builder.Append("- ");
                builder.Append(issue);
            }

            return builder.ToString();
        }
    }
}
