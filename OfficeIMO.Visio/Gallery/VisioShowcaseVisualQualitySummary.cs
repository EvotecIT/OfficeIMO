using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Summarizes deterministic visual-quality proof details for a generated showcase diagram.
    /// </summary>
    public sealed class VisioShowcaseVisualQualitySummary {
        internal static readonly VisioShowcaseVisualQualitySummary Empty = new(
            false,
            false,
            0,
            0,
            0,
            0,
            Array.Empty<string>());

        internal VisioShowcaseVisualQualitySummary(
            bool hasProof,
            bool isClean,
            int issueCount,
            int errorCount,
            int warningCount,
            int informationCount,
            IReadOnlyList<string> issueKinds) {
            HasProof = hasProof;
            IsClean = isClean;
            IssueCount = issueCount;
            ErrorCount = errorCount;
            WarningCount = warningCount;
            InformationCount = informationCount;
            IssueKinds = issueKinds;
        }

        /// <summary>Whether a visual-quality proof artifact was available and parsed.</summary>
        public bool HasProof { get; }

        /// <summary>Whether the parsed visual-quality proof reported the diagram as clean.</summary>
        public bool IsClean { get; }

        /// <summary>Total visual-quality issue count reported by the proof.</summary>
        public int IssueCount { get; }

        /// <summary>Error-severity issue count reported by the proof.</summary>
        public int ErrorCount { get; }

        /// <summary>Warning-severity issue count reported by the proof.</summary>
        public int WarningCount { get; }

        /// <summary>Information-severity issue count reported by the proof.</summary>
        public int InformationCount { get; }

        /// <summary>Distinct visual-quality issue kinds reported by the proof.</summary>
        public IReadOnlyList<string> IssueKinds { get; }
    }
}
