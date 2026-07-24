namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private enum FrameKind {
        Root,
        Italic,
        Bold,
        Strike,
        Highlight,
        Inserted,
        Superscript,
        Subscript,
    }

    private sealed class InlineFrame {
        public InlineFrame(FrameKind kind, char marker, int openLen, InlineSequence seq, int openIndex) {
            Kind = kind;
            Marker = marker;
            OpenLen = openLen;
            Seq = seq;
            OpenIndex = openIndex;
        }

        public FrameKind Kind { get; }
        public char Marker { get; }
        public int OpenLen { get; }
        public InlineSequence Seq { get; }
        public int OpenIndex { get; }
    }

    private sealed class EmphasisClosingRunIndex {
        private const byte HasRunOfAtLeastTwo = 1 << 2;
        private const byte SingleRunCountMask = 0x03;
        private readonly byte[] _asteriskSuffixSummaries;
        private readonly byte[] _underscoreSuffixSummaries;
        private readonly ClosingRunLengthIndex _asteriskOddRuns;
        private readonly ClosingRunLengthIndex _asteriskEvenRuns;
        private readonly ClosingRunLengthIndex _underscoreOddRuns;
        private readonly ClosingRunLengthIndex _underscoreEvenRuns;

        private EmphasisClosingRunIndex(
            byte[] asteriskSuffixSummaries,
            byte[] underscoreSuffixSummaries,
            ClosingRunLengthIndex asteriskOddRuns,
            ClosingRunLengthIndex asteriskEvenRuns,
            ClosingRunLengthIndex underscoreOddRuns,
            ClosingRunLengthIndex underscoreEvenRuns) {
            _asteriskSuffixSummaries = asteriskSuffixSummaries;
            _underscoreSuffixSummaries = underscoreSuffixSummaries;
            _asteriskOddRuns = asteriskOddRuns;
            _asteriskEvenRuns = asteriskEvenRuns;
            _underscoreOddRuns = underscoreOddRuns;
            _underscoreEvenRuns = underscoreEvenRuns;
        }

        internal static EmphasisClosingRunIndex Build(string text, bool cjkFriendlyEmphasis) {
            var asteriskSummaries = new byte[text.Length + 1];
            var underscoreSummaries = new byte[text.Length + 1];
            var asteriskOddRuns = new Dictionary<int, int>();
            var asteriskEvenRuns = new Dictionary<int, int>();
            var underscoreOddRuns = new Dictionary<int, int>();
            var underscoreEvenRuns = new Dictionary<int, int>();

            for (int index = text.Length - 1; index >= 0; index--) {
                asteriskSummaries[index] = asteriskSummaries[index + 1];
                underscoreSummaries[index] = underscoreSummaries[index + 1];

                char marker = text[index];
                if ((marker != '*' && marker != '_') || (index > 0 && text[index - 1] == marker)) {
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

                Dictionary<int, int> runStarts;
                if (marker == '*') {
                    runStarts = (runLength & 1) == 0 ? asteriskEvenRuns : asteriskOddRuns;
                } else {
                    runStarts = (runLength & 1) == 0 ? underscoreEvenRuns : underscoreOddRuns;
                }
                if (!runStarts.TryGetValue(runLength, out int currentMaximumStart) || index > currentMaximumStart) {
                    runStarts[runLength] = index;
                }

                byte[] summaries = marker == '*' ? asteriskSummaries : underscoreSummaries;
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
                ClosingRunLengthIndex.Create(asteriskOddRuns),
                ClosingRunLengthIndex.Create(asteriskEvenRuns),
                ClosingRunLengthIndex.Create(underscoreOddRuns),
                ClosingRunLengthIndex.Create(underscoreEvenRuns));
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

        private byte GetSummary(int start, char marker) {
            byte[] summaries = marker == '*' ? _asteriskSuffixSummaries : _underscoreSuffixSummaries;
            int index = Math.Max(0, Math.Min(start, summaries.Length - 1));
            return summaries[index];
        }

        private sealed class ClosingRunLengthIndex {
            private readonly int[] _runLengths;
            private readonly int[] _maximumStarts;
            private readonly int _leafOffset;

            private ClosingRunLengthIndex(int[] runLengths, int[] maximumStarts, int leafOffset) {
                _runLengths = runLengths;
                _maximumStarts = maximumStarts;
                _leafOffset = leafOffset;
            }

            internal static ClosingRunLengthIndex Create(Dictionary<int, int> maximumStartsByRunLength) {
                int[] runLengths = maximumStartsByRunLength.Keys.ToArray();
                Array.Sort(runLengths);

                int leafOffset = 1;
                while (leafOffset < runLengths.Length) {
                    leafOffset <<= 1;
                }

                var maximumStarts = new int[leafOffset * 2];
                for (int index = 0; index < maximumStarts.Length; index++) {
                    maximumStarts[index] = -1;
                }
                for (int index = 0; index < runLengths.Length; index++) {
                    maximumStarts[leafOffset + index] = maximumStartsByRunLength[runLengths[index]];
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
    }

    private static bool TryCloseFrame(
        Stack<InlineFrame> stack,
        char marker,
        int remaining,
        MarkdownInlineSourceMap? sourceMap,
        int closingIndex,
        out int consumed) {
        consumed = 0;
        if (stack == null || stack.Count <= 1) return false;
        var top = stack.Peek();
        if (top.Marker != marker) return false;

        // Close the innermost matching frame only; this avoids crossing.
        if (top.Kind == FrameKind.Italic && remaining >= 1) {
            stack.Pop();
            var node = new ItalicSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, 1);
            stack.Peek().Seq.AddRaw(node);
            consumed = 1;
            return true;
        }
        if (top.Kind == FrameKind.Bold && remaining >= 2) {
            stack.Pop();
            var node = new BoldSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, 2);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        if (top.Kind == FrameKind.Strike && remaining == top.OpenLen) {
            stack.Pop();
            var node = new StrikethroughSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, top.OpenLen);
            stack.Peek().Seq.AddRaw(node);
            consumed = top.OpenLen;
            return true;
        }
        if (top.Kind == FrameKind.Highlight && remaining >= 2) {
            stack.Pop();
            var node = new HighlightSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, 2);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        if (top.Kind == FrameKind.Inserted && remaining >= 2) {
            stack.Pop();
            var node = new InsertedSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, 2);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        if (top.Kind == FrameKind.Superscript && remaining >= 1) {
            stack.Pop();
            var node = new SuperscriptSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, 1);
            stack.Peek().Seq.AddRaw(node);
            consumed = 1;
            return true;
        }
        if (top.Kind == FrameKind.Subscript && remaining >= 1) {
            stack.Pop();
            var node = new SubscriptSequenceInline(top.Seq);
            SetFormattingMarkerSpans(node, sourceMap, marker, top.OpenIndex, top.OpenLen, closingIndex, 1);
            stack.Peek().Seq.AddRaw(node);
            consumed = 1;
            return true;
        }
        return false;
    }

    private static void SetFormattingMarkerSpans(
        MarkdownInline inline,
        MarkdownInlineSourceMap? sourceMap,
        char marker,
        int openingIndex,
        int openingLength,
        int closingIndex,
        int closingLength) {
        if (inline == null || sourceMap == null) {
            return;
        }

        MarkdownInlineSourceSpans.Set(
            inline,
            sourceMap.GetSpan(openingIndex, closingIndex + closingLength - openingIndex));
        MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
            inline,
            openingLength > 0 ? new string(marker, openingLength) : string.Empty,
            sourceMap.GetSpan(openingIndex, openingLength),
            closingLength > 0 ? new string(marker, closingLength) : string.Empty,
            sourceMap.GetSpan(closingIndex, closingLength));
    }

    private static bool ShouldTreatSingleMarkerAsLiteralInsideBold(string text, int start, char marker, int runLen, Stack<InlineFrame> stack, bool cjkFriendlyEmphasis) {
        if (runLen != 1) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Bold || top.Marker != marker || top.OpenLen != 2) return false;

        int nextDoubleClose = FindNextClosingDelimiterRunIndex(text, start + 1, marker, requiredRunLength: 2, cjkFriendlyEmphasis);
        if (nextDoubleClose >= 0) {
            int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, nextDoubleClose + 2, marker, requiredRunLength: 1, cjkFriendlyEmphasis);
            if (trailingSingleClose >= 0) return false;
        }

        int nextRun = FindNextDelimiterRunLength(text, start + 1, marker);
        return nextRun == 2;
    }

    private static bool ShouldTreatDelimiterRunAsLiteral(string text, int start, char marker, int runLen, Stack<InlineFrame> stack, bool splitDoubleRunIntoDualItalic, bool cjkFriendlyEmphasis, out int literalRunLength) {
        literalRunLength = 0;
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;

        var frames = stack.ToArray();
        if (frames.Length >= 2) {
            var parent = frames[1];
            // Keep the leading triple-delimiter path available for rebalancing into <em><strong>... later.
            if (parent.Kind == FrameKind.Bold && parent.Marker == marker && parent.OpenLen == 2 && parent.Seq.Nodes.Count == 0) return false;

            if (parent.Kind == FrameKind.Bold && parent.Marker == marker && parent.OpenLen == 2 && parent.Seq.Nodes.Count > 0) {
                int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1, cjkFriendlyEmphasis);
                if (trailingSingleClose >= 0) return false;
            }
        }

        if (splitDoubleRunIntoDualItalic) return false;

        int nextRunIndex = FindNextDelimiterRunIndex(text, start + 2, marker, out int nextRun);
        if (nextRun != 1) return false;

        GetDelimiterFlags(text, nextRunIndex, marker, nextRun, cjkFriendlyEmphasis, out bool nextCanOpen, out bool nextCanClose);
        if (nextCanOpen && !nextCanClose) return false;

        literalRunLength = 2;
        return true;
    }

    private static int FindNextDelimiterRunLength(string text, int start, char marker) {
        _ = FindNextDelimiterRunIndex(text, start, marker, out int runLength);
        return runLength;
    }

    private static int FindNextDelimiterRunIndex(string text, int start, char marker, out int runLength) {
        runLength = 0;
        if (string.IsNullOrEmpty(text)) return -1;
        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int run = 1;
            while (i + run < text.Length && text[i + run] == marker) run++;
            runLength = run;
            return i;
        }
        return -1;
    }

    private static bool TryRebalanceLeadingBoldInsideItalic(Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (remaining < 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (stack == null || stack.Count < 3) return false;

        var frames = stack.ToArray();
        var top = frames[0];
        var parent = frames[1];
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;
        if (parent.Kind != FrameKind.Bold || parent.Marker != marker || parent.OpenLen != 2) return false;
        if (parent.Seq.Nodes.Count != 0) return false;

        stack.Pop();
        stack.Pop();

        var italic = new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, parent.OpenIndex);
        italic.Seq.AddRaw(new BoldSequenceInline(top.Seq));
        stack.Push(italic);
        consumed = 2;
        return true;
    }

    private static bool TryRebalanceParentBoldWithInnerItalicIntoDualItalic(string text, int start, Stack<InlineFrame> stack, char marker, int remaining, bool cjkFriendlyEmphasis, out int consumed) {
        consumed = 0;
        if (remaining != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count < 3) return false;

        var frames = stack.ToArray();
        var top = frames[0];
        var parent = frames[1];
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;
        if (parent.Kind != FrameKind.Bold || parent.Marker != marker || parent.OpenLen != 2) return false;
        if (parent.Seq.Nodes.Count == 0) return false;

        int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1, cjkFriendlyEmphasis);
        if (trailingSingleClose < 0) return false;

        stack.Pop();
        stack.Pop();

        var middle = new InlineSequence { AutoSpacing = false };
        foreach (var node in parent.Seq.Nodes) {
            middle.AddRaw(node);
        }

        middle.AddRaw(new ItalicSequenceInline(top.Seq));

        var outer = new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, parent.OpenIndex);
        outer.Seq.AddRaw(new ItalicSequenceInline(middle));
        stack.Push(outer);
        consumed = 2;
        return true;
    }

    private static bool ShouldPreferInnerBold(Stack<InlineFrame> stack, char marker, int remaining, bool canOpen, bool canClose) {
        if (!canOpen || !canClose || remaining != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        return top.Marker == marker && top.Kind == FrameKind.Italic;
    }

    private static bool ShouldSplitDoubleUnderscoreToLiteralAndItalic(
        string text,
        int start,
        int runLen,
        bool canOpen,
        bool canClose,
        EmphasisClosingRunIndex? closingRuns) {
        if (!canOpen || canClose) return false;
        if (runLen != 2) return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (text[start] != '_') return false;
        if (closingRuns == null) return false;

        return !closingRuns.HasRunAtLeastTwo(start + 2, '_') &&
               closingRuns.CountSingleRuns(start + 2, '_') == 1;
    }

    private static bool ShouldSplitDoubleRunIntoRootDualItalic(
        string text,
        int start,
        char marker,
        int runLen,
        bool canOpen,
        bool canClose,
        Stack<InlineFrame> stack,
        EmphasisClosingRunIndex? closingRuns) {
        if (!canOpen || canClose) return false;
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count != 1) return false;
        if (closingRuns == null || closingRuns.HasRunAtLeastTwo(start + 2, marker)) return false;

        return closingRuns.CountSingleRuns(start + 2, marker) >= 2;
    }

    private static int GetLiteralPrefixLengthForOddCloser(
        string text,
        int start,
        char marker,
        int runLen,
        bool canOpen,
        bool canClose,
        EmphasisClosingRunIndex? closingRuns) {
        if (!canOpen || canClose) return 0;
        if (runLen < 2 || (runLen % 2) != 0) return 0;
        if (marker != '*' && marker != '_') return 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return 0;
        if (closingRuns == null) return 0;

        int suffixStart = start + runLen;
        int largestOddCloser = closingRuns.GetLargestRunAtMost(suffixStart, marker, runLen, odd: true);
        int largestEvenCloser = closingRuns.GetLargestRunAtMost(suffixStart, marker, runLen, odd: false);

        return largestOddCloser > largestEvenCloser ? runLen - largestOddCloser : 0;
    }

    private static bool HasFutureClosingDelimiterRun(string text, int start, char marker, int minimumRunLength, bool cjkFriendlyEmphasis) {
        if (string.IsNullOrEmpty(text)) return false;
        if (minimumRunLength <= 0) minimumRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, cjkFriendlyEmphasis, out _, out bool canClose);
            if (canClose && runLen >= minimumRunLength) return true;

            i += runLen - 1;
        }

        return false;
    }

    private static int CountFutureClosingDelimiterRuns(string text, int start, char marker, int requiredRunLength, int maximumCount, bool cjkFriendlyEmphasis) {
        if (string.IsNullOrEmpty(text)) return 0;
        if (requiredRunLength <= 0) requiredRunLength = 1;
        if (maximumCount <= 0) maximumCount = int.MaxValue;

        int count = 0;
        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, cjkFriendlyEmphasis, out _, out bool canClose);
            if (canClose && runLen == requiredRunLength) {
                count++;
                if (count >= maximumCount) return count;
            }

            i += runLen - 1;
        }

        return count;
    }

    private static bool ShouldTreatMixedSingleMarkerAsLiteral(string text, int start, char marker, int runLen, bool canOpen, bool canClose, Stack<InlineFrame> stack, bool cjkFriendlyEmphasis) {
        if (!canOpen || canClose) return false;
        if (runLen != 1) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if ((top.Kind != FrameKind.Italic && top.Kind != FrameKind.Bold) || (top.OpenLen != 1 && top.OpenLen != 2)) return false;
        if (top.Marker == marker) return false;

        int outerClose = FindNextClosingDelimiterRunIndex(text, start + 1, top.Marker, requiredRunLength: top.OpenLen, cjkFriendlyEmphasis);
        if (outerClose < 0) return false;

        int innerClose = FindNextClosingDelimiterIndex(text, start + 1, marker, minimumRunLength: 1, cjkFriendlyEmphasis);
        return innerClose < 0 || outerClose < innerClose;
    }

    private static bool ShouldTreatOppositeMarkerBeforeOuterCloseAsLiteral(string text, int start, char marker, int runLen, Stack<InlineFrame> stack, bool cjkFriendlyEmphasis) {
        if (runLen <= 0) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if ((top.Kind != FrameKind.Italic && top.Kind != FrameKind.Bold) || top.Marker == marker) return false;
        if (start + runLen >= text.Length) return false;

        char outerMarker = top.Marker;
        for (int i = 0; i < top.OpenLen; i++) {
            if (start + runLen + i >= text.Length || text[start + runLen + i] != outerMarker) {
                return false;
            }
        }

        GetDelimiterFlags(text, start + runLen, outerMarker, top.OpenLen, cjkFriendlyEmphasis, out _, out bool outerCanClose);
        return outerCanClose;
    }

    private static bool ShouldSplitDoubleRunIntoDualItalic(string text, int start, char marker, int runLen, Stack<InlineFrame> stack, bool cjkFriendlyEmphasis) {
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;

        int singleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1, cjkFriendlyEmphasis);
        if (singleClose < 0) return false;

        int immediateDoubleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 2, cjkFriendlyEmphasis);
        if (immediateDoubleClose >= 0 && immediateDoubleClose < singleClose) return false;

        if (HasOpeningDelimiterRunBefore(text, start + 2, singleClose, marker, requiredRunLength: 1, cjkFriendlyEmphasis)) return false;

        int doubleClose = FindNextClosingDelimiterRunIndex(text, singleClose + 1, marker, requiredRunLength: 2, cjkFriendlyEmphasis);
        if (doubleClose < 0) return false;

        int afterSingle = singleClose + 1;
        return afterSingle < text.Length && char.IsWhiteSpace(text[afterSingle]);
    }

    private static bool HasOpeningDelimiterRunBefore(string text, int start, int endExclusive, char marker, int requiredRunLength, bool cjkFriendlyEmphasis) {
        if (string.IsNullOrEmpty(text)) return false;
        if (requiredRunLength <= 0) requiredRunLength = 1;
        if (endExclusive <= start) return false;

        for (int i = Math.Max(0, start); i < Math.Min(text.Length, endExclusive); i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, cjkFriendlyEmphasis, out bool canOpen, out _);
            if (canOpen && runLen == requiredRunLength) return true;

            i += runLen - 1;
        }

        return false;
    }

    private static int FindNextClosingDelimiterIndex(string text, int start, char marker, int minimumRunLength, bool cjkFriendlyEmphasis) {
        if (string.IsNullOrEmpty(text)) return -1;
        if (minimumRunLength <= 0) minimumRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, cjkFriendlyEmphasis, out _, out bool canClose);
            if (canClose && runLen >= minimumRunLength) return i;

            i += runLen - 1;
        }

        return -1;
    }

    private static int FindNextClosingDelimiterRunIndex(string text, int start, char marker, int requiredRunLength, bool cjkFriendlyEmphasis) {
        if (string.IsNullOrEmpty(text)) return -1;
        if (requiredRunLength <= 0) requiredRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, cjkFriendlyEmphasis, out _, out bool canClose);
            if (canClose && runLen == requiredRunLength) return i;

            i += runLen - 1;
        }

        return -1;
    }

    private static void GetDelimiterFlags(string text, int start, char marker, int runLen, bool cjkFriendlyEmphasis, out bool canOpen, out bool canClose) {
        canOpen = false;
        canClose = false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return;

        char prev = start > 0 ? text[start - 1] : '\0';
        int nextIndex = start + runLen;
        char next = nextIndex < text.Length ? text[nextIndex] : '\0';

        bool prevWs = prev == '\0' || char.IsWhiteSpace(prev);
        bool nextWs = next == '\0' || char.IsWhiteSpace(next);
        bool prevPunct = prev != '\0' && IsPunctuationOrSymbol(prev);
        bool nextPunct = next != '\0' && IsPunctuationOrSymbol(next);
        bool prevCjk = IsCjkFriendlyEmphasisCharacter(prev);
        bool nextCjk = IsCjkFriendlyEmphasisCharacter(next);

        bool leftFlanking = !nextWs && (!nextPunct || prevWs || prevPunct || (cjkFriendlyEmphasis && marker == '*' && (prevCjk || nextCjk)));
        bool rightFlanking = !prevWs && (!prevPunct || nextWs || nextPunct || (cjkFriendlyEmphasis && marker == '*' && (prevCjk || nextCjk)));

        if (marker == '~') {
            // Pragmatic GFM-like strike and Markdig emphasis-extra subscript both hug non-whitespace text.
            canOpen = !nextWs;
            canClose = !prevWs;
            return;
        }

        if (marker == '=') {
            // Pragmatic mark/highlight handling: "==" opens/closes when it hugs non-whitespace text.
            canOpen = runLen >= 2 && !nextWs;
            canClose = runLen >= 2 && !prevWs;
            return;
        }

        if (marker == '+') {
            // Markdig emphasis extras: "++" opens/closes inserted text when it hugs non-whitespace text.
            canOpen = runLen >= 2 && !nextWs;
            canClose = runLen >= 2 && !prevWs;
            return;
        }

        if (marker == '^') {
            // Markdig emphasis extras: "^" opens/closes superscript when it hugs non-whitespace text.
            canOpen = !nextWs;
            canClose = !prevWs;
            return;
        }

        if (marker == '*') {
            canOpen = leftFlanking;
            canClose = rightFlanking;
            return;
        }

        // '_' is more restrictive (avoid intraword emphasis like foo_bar_baz).
        if (marker == '_') {
            canOpen = leftFlanking && (!rightFlanking || prevPunct || prevWs);
            canClose = rightFlanking && (!leftFlanking || nextPunct || nextWs);
            return;
        }
    }

    private static bool IsPunctuationOrSymbol(char c) => char.IsPunctuation(c) || char.IsSymbol(c);

    private static bool IsCjkFriendlyEmphasisCharacter(char c) =>
        c is >= '\u3000' and <= '\u303F'
            or >= '\u3040' and <= '\u30FF'
            or >= '\u31F0' and <= '\u31FF'
            or >= '\u3400' and <= '\u4DBF'
            or >= '\u4E00' and <= '\u9FFF'
            or >= '\uAC00' and <= '\uD7AF'
            or >= '\uFF00' and <= '\uFFEF';
}
