namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    // Answers which closing delimiter runs exist after a position, for one inline text. Every inline
    // parse of a text with '*' or '_' builds one, so the indexes and their buffers are reused per
    // thread; nested label parses take their own while their parent's is in use.
    private sealed class EmphasisClosingRunIndex : IDisposable {
        private const byte HasRunOfAtLeastTwo = 1 << 2;
        private const byte SingleRunCountMask = 0x03;
        private const int PoolSize = 4;
        private const int MaximumRetainedTextLength = 4096;

        [ThreadStatic] private static EmphasisClosingRunIndex?[]? t_pool;

        private byte[] _asteriskSuffixSummaries = Array.Empty<byte>();
        private byte[] _underscoreSuffixSummaries = Array.Empty<byte>();
        private bool _hasAsteriskSummaries;
        private bool _hasUnderscoreSummaries;
        private int _summaryLength;
        private readonly ClosingRunLengthIndex _asteriskOddRuns = new();
        private readonly ClosingRunLengthIndex _asteriskEvenRuns = new();
        private readonly ClosingRunLengthIndex _underscoreOddRuns = new();
        private readonly ClosingRunLengthIndex _underscoreEvenRuns = new();
        private readonly ClosingRunPositionIndex _asteriskPositions = new();
        private readonly ClosingRunPositionIndex _underscorePositions = new();
        private bool _disposed;

        private EmphasisClosingRunIndex() {
        }

        internal static EmphasisClosingRunIndex Build(string text, bool cjkFriendlyEmphasis) {
            var index = Rent();
            index.Fill(text, cjkFriendlyEmphasis);
            return index;
        }

        private static EmphasisClosingRunIndex Rent() {
            var pool = t_pool;
            if (pool != null) {
                for (int i = 0; i < pool.Length; i++) {
                    var pooled = pool[i];
                    if (pooled != null) {
                        pool[i] = null;
                        pooled._disposed = false;
                        return pooled;
                    }
                }
            }

            return new EmphasisClosingRunIndex();
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }
            _disposed = true;

            // Buffers sized for a very long text are left to the collector.
            if (_summaryLength > MaximumRetainedTextLength + 1) {
                return;
            }

            var pool = t_pool ??= new EmphasisClosingRunIndex?[PoolSize];
            for (int i = 0; i < pool.Length; i++) {
                if (pool[i] == null) {
                    pool[i] = this;
                    return;
                }
            }
        }

        private void Fill(string text, bool cjkFriendlyEmphasis) {
            int summaryLength = text.Length + 1;
            _summaryLength = summaryLength;
            _hasAsteriskSummaries = text.IndexOf('*') >= 0;
            _hasUnderscoreSummaries = text.IndexOf('_') >= 0;
            byte[] asteriskSummaries = _hasAsteriskSummaries ? EnsureLength(ref _asteriskSuffixSummaries, summaryLength) : _asteriskSuffixSummaries;
            byte[] underscoreSummaries = _hasUnderscoreSummaries ? EnsureLength(ref _underscoreSuffixSummaries, summaryLength) : _underscoreSuffixSummaries;
            if (_hasAsteriskSummaries) asteriskSummaries[text.Length] = 0;
            if (_hasUnderscoreSummaries) underscoreSummaries[text.Length] = 0;
            _asteriskOddRuns.Reset();
            _asteriskEvenRuns.Reset();
            _underscoreOddRuns.Reset();
            _underscoreEvenRuns.Reset();
            _asteriskPositions.Reset();
            _underscorePositions.Reset();

            for (int index = text.Length - 1; index >= 0; index--) {
                if (_hasAsteriskSummaries) {
                    asteriskSummaries[index] = asteriskSummaries[index + 1];
                }
                if (_hasUnderscoreSummaries) {
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
                    _asteriskPositions.AddDescending(index, runLength);
                    ((runLength & 1) == 0 ? _asteriskEvenRuns : _asteriskOddRuns).Record(runLength, index);
                } else {
                    _underscorePositions.AddDescending(index, runLength);
                    ((runLength & 1) == 0 ? _underscoreEvenRuns : _underscoreOddRuns).Record(runLength, index);
                }

                byte[] summaries = marker == '*' ? asteriskSummaries : underscoreSummaries;
                if (runLength >= 2) {
                    summaries[index] |= HasRunOfAtLeastTwo;
                } else {
                    int singleCount = Math.Min(2, (summaries[index] & SingleRunCountMask) + 1);
                    summaries[index] = (byte)((summaries[index] & ~SingleRunCountMask) | singleCount);
                }
            }

            _asteriskOddRuns.Complete();
            _asteriskEvenRuns.Complete();
            _underscoreOddRuns.Complete();
            _underscoreEvenRuns.Complete();
            _asteriskPositions.Complete();
            _underscorePositions.Complete();
        }

        // Returns a buffer of at least the length; a replaced buffer's contents are not kept.
        private static T[] EnsureLength<T>(ref T[] buffer, int length) {
            if (buffer.Length < length) {
                buffer = new T[Math.Max(length, Math.Min(buffer.Length * 2, int.MaxValue / 2))];
            }

            return buffer;
        }

        // Makes room for one more item after the first count items, keeping them.
        private static void EnsureRoomFor<T>(ref T[] buffer, int count) {
            if (count < buffer.Length) {
                return;
            }

            Array.Resize(ref buffer, Math.Max(4, buffer.Length * 2));
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
            bool present = marker == '*' ? _hasAsteriskSummaries : _hasUnderscoreSummaries;
            if (!present) {
                return 0;
            }

            byte[] summaries = marker == '*' ? _asteriskSuffixSummaries : _underscoreSuffixSummaries;
            int index = Math.Max(0, Math.Min(start, _summaryLength - 1));
            return summaries[index];
        }

        // The largest start of each closing run length, answering "the largest run of at most this
        // length that starts at or after a position" with a max segment tree over the sorted lengths.
        private sealed class ClosingRunLengthIndex {
            private int[] _runLengths = Array.Empty<int>();
            private int[] _recordedStarts = Array.Empty<int>();
            private int[] _maximumStarts = Array.Empty<int>();
            private int _count;
            private int _leafOffset = 1;

            internal void Reset() {
                _count = 0;
                _leafOffset = 1;
            }

            // Runs are recorded from the end of the text backwards, so the first start recorded
            // for a length is its largest.
            internal void Record(int runLength, int start) {
                for (int index = 0; index < _count; index++) {
                    if (_runLengths[index] == runLength) {
                        if (start > _recordedStarts[index]) {
                            _recordedStarts[index] = start;
                        }
                        return;
                    }
                }

                EnsureRoomFor(ref _runLengths, _count);
                EnsureRoomFor(ref _recordedStarts, _count);
                _runLengths[_count] = runLength;
                _recordedStarts[_count] = start;
                _count++;
            }

            internal void Complete() {
                if (_count == 0) {
                    return;
                }

                // Lengths are distinct, so sorting them with their starts is deterministic.
                Array.Sort(_runLengths, _recordedStarts, 0, _count);

                int leafOffset = 1;
                while (leafOffset < _count) {
                    leafOffset <<= 1;
                }

                _leafOffset = leafOffset;
                int treeLength = leafOffset * 2;
                EnsureLength(ref _maximumStarts, treeLength);
                for (int index = 0; index < treeLength; index++) {
                    _maximumStarts[index] = -1;
                }
                for (int index = 0; index < _count; index++) {
                    _maximumStarts[leafOffset + index] = _recordedStarts[index];
                }
                for (int index = leafOffset - 1; index > 0; index--) {
                    _maximumStarts[index] = Math.Max(_maximumStarts[index * 2], _maximumStarts[(index * 2) + 1]);
                }
            }

            internal int GetLargestRunAtMost(int maximumRunLength, int minimumStart) {
                if (_count == 0 || maximumRunLength <= 0) {
                    return 0;
                }

                int maximumIndex = Array.BinarySearch(_runLengths, 0, _count, maximumRunLength);
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
                return matchingIndex >= 0 && matchingIndex < _count
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

        // Starts of closing runs in text order: all of them, and those of exactly one and two markers.
        private sealed class ClosingRunPositionIndex {
            private int[] _allStarts = Array.Empty<int>();
            private int[] _singleStarts = Array.Empty<int>();
            private int[] _doubleStarts = Array.Empty<int>();
            private int _allCount;
            private int _singleCount;
            private int _doubleCount;

            internal void Reset() {
                _allCount = 0;
                _singleCount = 0;
                _doubleCount = 0;
            }

            // Runs arrive from the end of the text backwards and are put in text order by Complete.
            internal void AddDescending(int start, int length) {
                Append(ref _allStarts, ref _allCount, start);
                if (length == 1) {
                    Append(ref _singleStarts, ref _singleCount, start);
                } else if (length == 2) {
                    Append(ref _doubleStarts, ref _doubleCount, start);
                }
            }

            private static void Append(ref int[] starts, ref int count, int start) {
                EnsureRoomFor(ref starts, count);
                starts[count++] = start;
            }

            internal void Complete() {
                Array.Reverse(_allStarts, 0, _allCount);
                Array.Reverse(_singleStarts, 0, _singleCount);
                Array.Reverse(_doubleStarts, 0, _doubleCount);
            }

            internal int FindFirst(int minimumStart, int exactRunLength) {
                int[] starts;
                int count;
                switch (exactRunLength) {
                    case 0:
                        starts = _allStarts;
                        count = _allCount;
                        break;
                    case 1:
                        starts = _singleStarts;
                        count = _singleCount;
                        break;
                    case 2:
                        starts = _doubleStarts;
                        count = _doubleCount;
                        break;
                    default:
                        return -1;
                }

                int index = Array.BinarySearch(starts, 0, count, minimumStart);
                if (index < 0) index = ~index;
                return index < count ? starts[index] : -1;
            }
        }
    }

}
