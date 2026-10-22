#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private sealed class EmphasisClosingRunIndex : IDisposable {
        private const byte HasRunOfAtLeastTwo = 1 << 2;
        private const byte SingleRunCountMask = 0x03;
        private readonly byte[]? _asteriskSuffixSummaries;
        private readonly byte[]? _underscoreSuffixSummaries;
        private readonly int _summaryLength;
        private readonly ClosingRunLengthIndex _asteriskOddRuns;
        private readonly ClosingRunLengthIndex _asteriskEvenRuns;
        private readonly ClosingRunLengthIndex _underscoreOddRuns;
        private readonly ClosingRunLengthIndex _underscoreEvenRuns;
        private readonly ClosingRunPositionIndex _asteriskPositions;
        private readonly ClosingRunPositionIndex _underscorePositions;
        private bool _disposed;

        private EmphasisClosingRunIndex(
            byte[]? asteriskSuffixSummaries,
            byte[]? underscoreSuffixSummaries,
            int summaryLength,
            ClosingRunLengthIndex asteriskOddRuns,
            ClosingRunLengthIndex asteriskEvenRuns,
            ClosingRunLengthIndex underscoreOddRuns,
            ClosingRunLengthIndex underscoreEvenRuns,
            ClosingRunPositionIndex asteriskPositions,
            ClosingRunPositionIndex underscorePositions) {
            _asteriskSuffixSummaries = asteriskSuffixSummaries;
            _underscoreSuffixSummaries = underscoreSuffixSummaries;
            _summaryLength = summaryLength;
            _asteriskOddRuns = asteriskOddRuns;
            _asteriskEvenRuns = asteriskEvenRuns;
            _underscoreOddRuns = underscoreOddRuns;
            _underscoreEvenRuns = underscoreEvenRuns;
            _asteriskPositions = asteriskPositions;
            _underscorePositions = underscorePositions;
        }

        internal static EmphasisClosingRunIndex Build(string text, bool cjkFriendlyEmphasis) {
            int summaryLength = text.Length + 1;
            byte[]? asteriskSummaries = text.IndexOf('*') >= 0
                ? RentSummaryBuffer(summaryLength)
                : null;
            byte[]? underscoreSummaries = text.IndexOf('_') >= 0
                ? RentSummaryBuffer(summaryLength)
                : null;
            if (asteriskSummaries != null) asteriskSummaries[text.Length] = 0;
            if (underscoreSummaries != null) underscoreSummaries[text.Length] = 0;
            List<ClosingRunLengthSummary>? asteriskOddRuns = null;
            List<ClosingRunLengthSummary>? asteriskEvenRuns = null;
            List<ClosingRunLengthSummary>? underscoreOddRuns = null;
            List<ClosingRunLengthSummary>? underscoreEvenRuns = null;
            List<(int Start, int Length)>? asteriskPositions = null;
            List<(int Start, int Length)>? underscorePositions = null;

            for (int index = text.Length - 1; index >= 0; index--) {
                if (asteriskSummaries != null) {
                    asteriskSummaries[index] = asteriskSummaries[index + 1];
                }
                if (underscoreSummaries != null) {
                    underscoreSummaries[index] = underscoreSummaries[index + 1];
                }

                char marker = text[index];
                if ((marker != '*' && marker != '_')
                    || IsEscapedEmphasisMarker(text, index)
                    || (index > 0 && text[index - 1] == marker && !IsEscapedEmphasisMarker(text, index - 1))) {
                    continue;
                }

                int runLength = 1;
                while (index + runLength < text.Length && text[index + runLength] == marker) {
                    runLength++;
                }

                GetDelimiterFlags(text, index, marker, runLength, cjkFriendlyEmphasis, out _, out bool canClose);
                if (!canClose) {
                    continue;
                }

                if (marker == '*') {
                    (asteriskPositions ??= new List<(int Start, int Length)>()).Add((index, runLength));
                    if ((runLength & 1) == 0) {
                        RecordRunLength(ref asteriskEvenRuns, runLength, index);
                    } else {
                        RecordRunLength(ref asteriskOddRuns, runLength, index);
                    }
                } else {
                    (underscorePositions ??= new List<(int Start, int Length)>()).Add((index, runLength));
                    if ((runLength & 1) == 0) {
                        RecordRunLength(ref underscoreEvenRuns, runLength, index);
                    } else {
                        RecordRunLength(ref underscoreOddRuns, runLength, index);
                    }
                }

                byte[] summaries = marker == '*' ? asteriskSummaries! : underscoreSummaries!;
                if (runLength >= 2) {
                    summaries[index] |= HasRunOfAtLeastTwo;
                } else {
                    int singleCount = Math.Min(2, (summaries[index] & SingleRunCountMask) + 1);
                    summaries[index] = (byte)((summaries[index] & ~SingleRunCountMask) | singleCount);
                }
            }

            return new EmphasisClosingRunIndex(
                asteriskSummaries,
                underscoreSummaries,
                summaryLength,
                ClosingRunLengthIndex.Create(asteriskOddRuns),
                ClosingRunLengthIndex.Create(asteriskEvenRuns),
                ClosingRunLengthIndex.Create(underscoreOddRuns),
                ClosingRunLengthIndex.Create(underscoreEvenRuns),
                ClosingRunPositionIndex.Create(asteriskPositions),
                ClosingRunPositionIndex.Create(underscorePositions));
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }
            _disposed = true;

            if (_asteriskSuffixSummaries != null) {
                ReturnSummaryBuffer(_asteriskSuffixSummaries);
            }
            if (_underscoreSuffixSummaries != null) {
                ReturnSummaryBuffer(_underscoreSuffixSummaries);
            }
        }

        private static byte[] RentSummaryBuffer(int minimumLength) {
#if NET8_0_OR_GREATER
            return ArrayPool<byte>.Shared.Rent(minimumLength);
#else
            return new byte[minimumLength];
#endif
        }

        private static void ReturnSummaryBuffer(byte[] buffer) {
#if NET8_0_OR_GREATER
            ArrayPool<byte>.Shared.Return(buffer, clearArray: false);
#endif
        }

        private static void RecordRunLength(
            ref List<ClosingRunLengthSummary>? runs,
            int runLength,
            int start) {
            runs ??= new List<ClosingRunLengthSummary>(2);
            for (int index = 0; index < runs.Count; index++) {
                if (runs[index].RunLength != runLength) {
                    continue;
                }

                if (start > runs[index].MaximumStart) {
                    runs[index] = new ClosingRunLengthSummary(runLength, start);
                }
                return;
            }

            runs.Add(new ClosingRunLengthSummary(runLength, start));
        }

        internal bool HasRunAtLeastTwo(int start, char marker) {
            byte summary = GetSummary(start, marker);
            return (summary & HasRunOfAtLeastTwo) != 0;
        }

        internal int CountSingleRuns(int start, char marker) => GetSummary(start, marker) & SingleRunCountMask;

        internal int GetLargestRunAtMost(int start, char marker, int maximumRunLength, bool odd) {
            ClosingRunLengthIndex index;
            if (marker == '*') {
                index = odd ? _asteriskOddRuns : _asteriskEvenRuns;
            } else {
                index = odd ? _underscoreOddRuns : _underscoreEvenRuns;
            }

            return index.GetLargestRunAtMost(maximumRunLength, start);
        }

        internal int FindFirstRun(int start, char marker, int exactRunLength) {
            ClosingRunPositionIndex index = marker == '*'
                ? _asteriskPositions
                : _underscorePositions;
            return index.FindFirst(start, exactRunLength);
        }

        internal int FindFirstRun(int start, char marker) {
            ClosingRunPositionIndex index = marker == '*'
                ? _asteriskPositions
                : _underscorePositions;
            return index.FindFirst(start, exactRunLength: 0);
        }

        private byte GetSummary(int start, char marker) {
            byte[]? summaries = marker == '*' ? _asteriskSuffixSummaries : _underscoreSuffixSummaries;
            if (summaries == null) {
                return 0;
            }
            int index = Math.Max(0, Math.Min(start, _summaryLength - 1));
            return summaries[index];
        }

        private readonly record struct ClosingRunLengthSummary(int RunLength, int MaximumStart);

        private sealed class ClosingRunLengthIndex {
            private static readonly ClosingRunLengthIndex Empty = new(
                Array.Empty<int>(),
                Array.Empty<int>(),
                leafOffset: 1);
            private readonly int[] _runLengths;
            private readonly int[] _maximumStarts;
            private readonly int _leafOffset;

            private ClosingRunLengthIndex(int[] runLengths, int[] maximumStarts, int leafOffset) {
                _runLengths = runLengths;
                _maximumStarts = maximumStarts;
                _leafOffset = leafOffset;
            }

            internal static ClosingRunLengthIndex Create(List<ClosingRunLengthSummary>? runSummaries) {
                if (runSummaries == null || runSummaries.Count == 0) {
                    return Empty;
                }

                runSummaries.Sort(static (left, right) => left.RunLength.CompareTo(right.RunLength));
                var runLengths = new int[runSummaries.Count];

                int leafOffset = 1;
                while (leafOffset < runSummaries.Count) {
                    leafOffset <<= 1;
                }

                var maximumStarts = new int[leafOffset * 2];
                for (int index = 0; index < maximumStarts.Length; index++) {
                    maximumStarts[index] = -1;
                }
                for (int index = 0; index < runSummaries.Count; index++) {
                    runLengths[index] = runSummaries[index].RunLength;
                    maximumStarts[leafOffset + index] = runSummaries[index].MaximumStart;
                }
                for (int index = leafOffset - 1; index > 0; index--) {
                    maximumStarts[index] = Math.Max(maximumStarts[index * 2], maximumStarts[(index * 2) + 1]);
                }

                return new ClosingRunLengthIndex(runLengths, maximumStarts, leafOffset);
            }

            internal int GetLargestRunAtMost(int maximumRunLength, int minimumStart) {
                if (_runLengths.Length == 0 || maximumRunLength <= 0) {
                    return 0;
                }

                int maximumIndex = Array.BinarySearch(_runLengths, maximumRunLength);
                if (maximumIndex < 0) {
                    maximumIndex = ~maximumIndex - 1;
                }
                if (maximumIndex < 0) {
                    return 0;
                }

                int matchingIndex = FindRightmostIndex(
                    node: 1,
                    segmentStart: 0,
                    segmentEnd: _leafOffset - 1,
                    queryEnd: maximumIndex,
                    minimumStart);
                return matchingIndex >= 0 && matchingIndex < _runLengths.Length
                    ? _runLengths[matchingIndex]
                    : 0;
            }

            private int FindRightmostIndex(int node, int segmentStart, int segmentEnd, int queryEnd, int minimumStart) {
                if (segmentStart > queryEnd || _maximumStarts[node] < minimumStart) {
                    return -1;
                }
                if (segmentStart == segmentEnd) {
                    return segmentStart;
                }

                int middle = segmentStart + ((segmentEnd - segmentStart) / 2);
                int right = FindRightmostIndex((node * 2) + 1, middle + 1, segmentEnd, queryEnd, minimumStart);
                return right >= 0
                    ? right
                    : FindRightmostIndex(node * 2, segmentStart, middle, queryEnd, minimumStart);
            }
        }

        private sealed class ClosingRunPositionIndex {
            private static readonly ClosingRunPositionIndex Empty = new(
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>());
            private readonly int[] _allStarts;
            private readonly int[] _singleStarts;
            private readonly int[] _doubleStarts;

            private ClosingRunPositionIndex(
                int[] allStarts,
                int[] singleStarts,
                int[] doubleStarts) {
                _allStarts = allStarts;
                _singleStarts = singleStarts;
                _doubleStarts = doubleStarts;
            }

            internal static ClosingRunPositionIndex Create(
                List<(int Start, int Length)>? descendingRuns) {
                if (descendingRuns == null || descendingRuns.Count == 0) {
                    return Empty;
                }

                int singleCount = 0;
                int doubleCount = 0;
                for (int index = 0; index < descendingRuns.Count; index++) {
                    if (descendingRuns[index].Length == 1) singleCount++;
                    else if (descendingRuns[index].Length == 2) doubleCount++;
                }

                var all = new int[descendingRuns.Count];
                var singles = new int[singleCount];
                var doubles = new int[doubleCount];
                int allIndex = 0;
                int singleIndex = 0;
                int doubleIndex = 0;
                for (int index = descendingRuns.Count - 1; index >= 0; index--) {
                    (int start, int length) = descendingRuns[index];
                    all[allIndex++] = start;
                    if (length == 1) singles[singleIndex++] = start;
                    else if (length == 2) doubles[doubleIndex++] = start;
                }
                return new ClosingRunPositionIndex(all, singles, doubles);
            }

            internal int FindFirst(int minimumStart, int exactRunLength) {
                int[] starts = exactRunLength switch {
                    0 => _allStarts,
                    1 => _singleStarts,
                    2 => _doubleStarts,
                    _ => Array.Empty<int>()
                };
                int index = Array.BinarySearch(starts, minimumStart);
                if (index < 0) index = ~index;
                return index < starts.Length ? starts[index] : -1;
            }
        }
    }

}
