using System;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static double GetContainmentAwareTextSimilarity(string source, string target) =>
            GetContainmentAwareTextSimilarity(source, target, null);

        private static double GetContainmentAwareTextSimilarity(
            string source,
            string target,
            ComparisonWorkBudget? comparisonWorkBudget) {
            long cellCount = ((long)source.Length + 1) * ((long)target.Length + 1);
            if (source.Length > 0 &&
                target.Length > 0 &&
                cellCount <= LcsCellLimit &&
                (comparisonWorkBudget == null || comparisonWorkBudget.TryConsume(Math.Max(source.Length, target.Length))) &&
                (source.IndexOf(target, StringComparison.Ordinal) >= 0 ||
                 target.IndexOf(source, StringComparison.Ordinal) >= 0)) {
                return 0.75 + (0.25 * Math.Min(source.Length, target.Length) / Math.Max(source.Length, target.Length));
            }

            return GetTextSimilarity(source, target, comparisonWorkBudget);
        }

        private static double GetTextSimilarity(string source, string target) =>
            GetTextSimilarity(source, target, null);

        private static double GetTextSimilarity(
            string source,
            string target,
            ComparisonWorkBudget? comparisonWorkBudget) {
            if (source.Length == 0 || target.Length == 0) {
                return source.Length == target.Length ? 1D : 0D;
            }

            long cellCount = ((long)source.Length + 1) * ((long)target.Length + 1);
            if (cellCount > LcsCellLimit) {
                return GetBoundedTextSimilarity(source, target, comparisonWorkBudget);
            }

            if (comparisonWorkBudget != null && !comparisonWorkBudget.TryConsume(cellCount)) {
                return 0D;
            }

            if (string.Equals(source, target, StringComparison.Ordinal)) {
                return 1D;
            }

            return (double)GetCommonSubsequenceLength(source, target) / Math.Max(source.Length, target.Length);
        }

        private static bool AreComparisonTextEqual(string? source, string? target, WordComparisonOptions options) =>
            string.Equals(NormalizeComparisonText(source ?? string.Empty, options), NormalizeComparisonText(target ?? string.Empty, options), StringComparison.Ordinal);

        private static string NormalizeComparisonText(string value, WordComparisonOptions options) {
            string normalized = value ?? string.Empty;
            if (options.IgnoreWhitespace) {
                normalized = string.Join(" ", normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            }

            if (options.IgnoreCase) {
                normalized = normalized.ToUpperInvariant();
            }

            return normalized;
        }

        private static double GetBoundedTextSimilarity(string source, string target) =>
            GetBoundedTextSimilarity(source, target, null);

        private static double GetBoundedTextSimilarity(
            string source,
            string target,
            ComparisonWorkBudget? comparisonWorkBudget) {
            int prefixLength = 0;
            int maxPrefixLength = Math.Min(source.Length, target.Length);
            int maximumEdgeLength = Math.Min(MaxBoundedTextSimilaritySamples, maxPrefixLength);
            long boundedWorkUnits = (maximumEdgeLength * 2L) + Math.Min(MaxBoundedTextSimilaritySamples, maxPrefixLength);
            if (comparisonWorkBudget != null && !comparisonWorkBudget.TryConsume(boundedWorkUnits)) {
                return 0D;
            }

            while (prefixLength < maximumEdgeLength && source[prefixLength] == target[prefixLength]) {
                prefixLength++;
            }

            int suffixLength = 0;
            int sourceSuffixIndex = source.Length - 1;
            int targetSuffixIndex = target.Length - 1;
            while (suffixLength < maximumEdgeLength &&
                   sourceSuffixIndex >= prefixLength &&
                   targetSuffixIndex >= prefixLength &&
                   source[sourceSuffixIndex] == target[targetSuffixIndex]) {
                suffixLength++;
                sourceSuffixIndex--;
                targetSuffixIndex--;
            }

            double edgeSimilarity = (double)(prefixLength + suffixLength) / Math.Max(source.Length, target.Length);
            int sampleCount = Math.Min(MaxBoundedTextSimilaritySamples, Math.Min(source.Length, target.Length));
            if (sampleCount <= 1) {
                return edgeSimilarity;
            }

            int matchingSamples = 0;
            for (int sample = 0; sample < sampleCount; sample++) {
                int sourceIndex = (int)((long)sample * (source.Length - 1) / (sampleCount - 1));
                int targetIndex = (int)((long)sample * (target.Length - 1) / (sampleCount - 1));
                if (source[sourceIndex] == target[targetIndex]) {
                    matchingSamples++;
                }
            }

            double lengthRatio = (double)Math.Min(source.Length, target.Length) / Math.Max(source.Length, target.Length);
            double sampledSimilarity = (double)matchingSamples / sampleCount * lengthRatio;
            double anchorSimilarity = GetBoundedAnchorTextSimilarity(source, target, lengthRatio, comparisonWorkBudget);
            return Math.Max(edgeSimilarity, Math.Max(sampledSimilarity, anchorSimilarity));
        }

        private static double GetBoundedAnchorTextSimilarity(
            string source,
            string target,
            double lengthRatio,
            ComparisonWorkBudget? comparisonWorkBudget) {
            string anchors = source.Length <= target.Length ? source : target;
            string searchable = source.Length <= target.Length ? target : source;
            if (anchors.Length == 0) return 0D;

            int anchorLength = Math.Min(BoundedTextSimilarityAnchorLength, anchors.Length);
            int maximumStart = anchors.Length - anchorLength;
            int maximumSearchStart = searchable.Length - anchorLength;
            int sampleCount = Math.Min(MaxBoundedTextSimilarityAnchors, maximumStart + 1);
            int matchingAnchors = 0;
            for (int sample = 0; sample < sampleCount; sample++) {
                int anchorStart = sampleCount == 1
                    ? 0
                    : (int)((long)sample * maximumStart / (sampleCount - 1));
                bool matched;
                if (comparisonWorkBudget != null) {
                    if (!comparisonWorkBudget.TryConsume(Math.Max(searchable.Length, MaxBoundedTextSimilaritySamples))) {
                        break;
                    }

                    string anchor = anchors.Substring(anchorStart, anchorLength);
                    matched = searchable.IndexOf(anchor, StringComparison.Ordinal) >= 0;
                } else {
                    int searchCenter = maximumStart == 0
                        ? 0
                        : (int)((long)anchorStart * maximumSearchStart / maximumStart);
                    matched = ContainsAnchorWithinBoundedOffset(
                        anchors,
                        anchorStart,
                        searchable,
                        searchCenter,
                        anchorLength,
                        maximumSearchStart);
                }

                if (matched) {
                    matchingAnchors++;
                }
            }

            return (double)matchingAnchors / sampleCount * lengthRatio;
        }

        private static bool ContainsAnchorWithinBoundedOffset(
            string anchors,
            int anchorStart,
            string searchable,
            int searchCenter,
            int anchorLength,
            int maximumSearchStart) {
            int searchStart = Math.Max(0, searchCenter - MaxBoundedTextSimilaritySamples);
            int searchEnd = Math.Min(maximumSearchStart, searchCenter + MaxBoundedTextSimilaritySamples);
            for (int candidateStart = searchStart; candidateStart <= searchEnd; candidateStart++) {
                int offset = 0;
                while (offset < anchorLength &&
                       anchors[anchorStart + offset] == searchable[candidateStart + offset]) {
                    offset++;
                }

                if (offset == anchorLength) {
                    return true;
                }
            }

            return false;
        }

        private static int GetCommonSubsequenceLength(string source, string target) {
            int[,] lengths = new int[source.Length + 1, target.Length + 1];

            for (int sourceIndex = source.Length - 1; sourceIndex >= 0; sourceIndex--) {
                for (int targetIndex = target.Length - 1; targetIndex >= 0; targetIndex--) {
                    lengths[sourceIndex, targetIndex] = source[sourceIndex] == target[targetIndex]
                        ? lengths[sourceIndex + 1, targetIndex + 1] + 1
                        : Math.Max(lengths[sourceIndex + 1, targetIndex], lengths[sourceIndex, targetIndex + 1]);
                }
            }

            return lengths[0, 0];
        }
    }
}
