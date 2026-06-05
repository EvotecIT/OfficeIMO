using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OfficeIMO.Visio {
    public sealed partial class VisioShowcaseSummary {
        private static VisioShowcaseVisualQualitySummary CreateVisualQualitySummary(string showcasePath, IReadOnlyList<VisioShowcaseArtifact> proofs) {
            VisioShowcaseArtifact? proof = proofs.FirstOrDefault(item => item.Kind == VisioShowcaseArtifactKind.VisualQuality);
            if (proof == null) {
                return VisioShowcaseVisualQualitySummary.Empty;
            }

            string proofPath = Path.Combine(showcasePath, proof.RelativePath.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(proofPath)) {
                return VisioShowcaseVisualQualitySummary.Empty;
            }

            bool isClean = false;
            int issueCount = 0;
            int errorCount = 0;
            int warningCount = 0;
            int informationCount = 0;
            Dictionary<string, string> issueKinds = new(StringComparer.OrdinalIgnoreCase);

            foreach (string line in File.ReadLines(proofPath)) {
                if (TryReadProofValue(line, "quality.isClean=", out string cleanValue)) {
                    isClean = string.Equals(cleanValue, "true", StringComparison.OrdinalIgnoreCase);
                } else if (TryReadNonNegativeInt(line, "quality.issueCount=", out int parsedIssueCount)) {
                    issueCount = parsedIssueCount;
                } else if (TryReadNonNegativeInt(line, "quality.errorCount=", out int parsedErrorCount)) {
                    errorCount = parsedErrorCount;
                } else if (TryReadNonNegativeInt(line, "quality.warningCount=", out int parsedWarningCount)) {
                    warningCount = parsedWarningCount;
                } else if (TryReadNonNegativeInt(line, "quality.informationCount=", out int parsedInformationCount)) {
                    informationCount = parsedInformationCount;
                } else if (TryReadProofValue(line, "quality.issueKinds=", out string parsedIssueKinds)) {
                    AddCsvValues(issueKinds, parsedIssueKinds);
                }
            }

            return new VisioShowcaseVisualQualitySummary(
                true,
                isClean,
                issueCount,
                errorCount,
                warningCount,
                informationCount,
                ToSortedReadOnlyList(issueKinds));
        }

        private static bool TryReadNonNegativeInt(string line, string prefix, out int value) {
            if (TryReadProofValue(line, prefix, out string rawValue) &&
                int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) &&
                value >= 0) {
                return true;
            }

            value = 0;
            return false;
        }
    }
}
