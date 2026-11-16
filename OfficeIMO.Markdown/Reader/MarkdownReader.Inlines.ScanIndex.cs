namespace OfficeIMO.Markdown;

// The inline parser asks the same forward questions ("where does this bracket close?", "where is
// the next backtick run of this length?") from many start positions in one text. Answering each
// question with a fresh scan to the end of the text made malformed input quadratic, so while an
// inline parse is running the answers for a text are computed in one pass and reused.
internal static partial class MarkdownReader {
    private static class InlineScanIndexCache {
        // Nested label and wrapper parses use their own texts while their parent's indexes are still
        // needed, so the most recently used entries stay at the front and the oldest one is dropped.
        private const int Capacity = 4;

        [ThreadStatic] private static int t_scopeDepth;
        [ThreadStatic] private static BracketMatchIndex?[]? t_bracketIndexes;
        [ThreadStatic] private static BacktickRunIndex?[]? t_backtickIndexes;
        [ThreadStatic] private static ForwardSearch?[]? t_forwardSearches;
        [ThreadStatic] private static WrapperIndexEntry?[]? t_wrapperIndexes;

        internal static Scope Enter() {
            t_scopeDepth++;
            return new Scope();
        }

        internal readonly struct Scope : IDisposable {
            public void Dispose() {
                if (--t_scopeDepth > 0) {
                    return;
                }

                // Indexes hold the paragraph text, so they don't outlive the parse that built them.
                // The slot arrays stay for the next parse on this thread.
                t_scopeDepth = 0;
                Clear(t_bracketIndexes);
                Clear(t_backtickIndexes);
                Clear(t_forwardSearches);
                Clear(t_wrapperIndexes);
                if (t_lineVerdict != null) {
                    t_lineVerdict.Text = null;
                }
            }
        }

        internal static BracketMatchIndex? GetBracketIndex(
            string text,
            MarkdownLoadOptions? options,
            InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches) {
            if (t_scopeDepth == 0) {
                return null;
            }

            bool inlineHtml = options?.InlineHtml != false;
            var entries = t_bracketIndexes ??= new BracketMatchIndex?[Capacity];
            for (int i = 0; i < entries.Length; i++) {
                var entry = entries[i];
                if (entry != null && entry.Matches(text, inlineHtml, inlineHtmlWrapperMatches)) {
                    MoveToFront(entries, i);
                    return entry;
                }
            }

            var created = BracketMatchIndex.Build(text, options, inlineHtml, inlineHtmlWrapperMatches);
            InsertAtFront(entries, created);
            return created;
        }

        internal static BacktickRunIndex? GetBacktickIndex(string text) {
            if (t_scopeDepth == 0) {
                return null;
            }

            var entries = t_backtickIndexes ??= new BacktickRunIndex?[Capacity];
            for (int i = 0; i < entries.Length; i++) {
                var entry = entries[i];
                if (entry != null && ReferenceEquals(entry.Text, text)) {
                    MoveToFront(entries, i);
                    return entry;
                }
            }

            var created = BacktickRunIndex.Build(text);
            InsertAtFront(entries, created);
            return created;
        }

        private sealed class WrapperIndexEntry {
            internal WrapperIndexEntry(string text, InlineHtmlWrapperMatchIndex index) {
                Text = text;
                Index = index;
            }

            internal string Text { get; }
            internal InlineHtmlWrapperMatchIndex Index { get; }
        }

        // Inline HTML wrapper pairs of a whole text, for callers that need them although the parse
        // they are part of was given none.
        internal static InlineHtmlWrapperMatchIndex GetWrapperIndex(string text) {
            if (t_scopeDepth == 0) {
                return BuildInlineHtmlWrapperMatchIndex(text);
            }

            var entries = t_wrapperIndexes ??= new WrapperIndexEntry?[Capacity];
            for (int i = 0; i < entries.Length; i++) {
                var entry = entries[i];
                if (entry != null && ReferenceEquals(entry.Text, text)) {
                    MoveToFront(entries, i);
                    return entry.Index;
                }
            }

            var created = new WrapperIndexEntry(text, BuildInlineHtmlWrapperMatchIndex(text));
            InsertAtFront(entries, created);
            return created.Index;
        }

        private sealed class ForwardSearch {
            internal ForwardSearch(string text, string value) {
                Text = text;
                Value = value;
            }

            internal string Text { get; }
            internal string Value { get; }
            internal int From { get; set; }
            internal int Result { get; set; }
        }

        // Ordinal IndexOf that reuses the previous answer for the same text and value: when no match
        // starts in [From, Result), every start inside that range has the same answer.
        internal static int IndexOf(string text, string value, int startIndex) {
            if (t_scopeDepth == 0) {
                return text.IndexOf(value, startIndex, StringComparison.Ordinal);
            }

            var entries = t_forwardSearches ??= new ForwardSearch?[Capacity];
            ForwardSearch? entry = null;
            for (int i = 0; i < entries.Length; i++) {
                var candidate = entries[i];
                if (candidate != null && ReferenceEquals(candidate.Text, text) && ReferenceEquals(candidate.Value, value)) {
                    MoveToFront(entries, i);
                    entry = candidate;
                    break;
                }
            }

            if (entry != null
                && startIndex >= entry.From
                && (entry.Result < 0 || startIndex <= entry.Result)) {
                return entry.Result;
            }

            int result = text.IndexOf(value, startIndex, StringComparison.Ordinal);
            if (entry == null) {
                entry = new ForwardSearch(text, value);
                InsertAtFront(entries, entry);
            }

            entry.From = startIndex;
            entry.Result = result;
            return result;
        }

        [ThreadStatic] private static LineVerdict? t_lineVerdict;

        // One verdict per thread, overwritten by each new line instead of allocated per line.
        private sealed class LineVerdict {
            internal string? Text;
            internal int LineStart;
            internal int LineEnd;
            internal bool Verdict;
        }

        // A position between a line's start and its terminating '\n' (or the text end) has the
        // same line bounds, so a verdict about the line can be reused for any of them.
        internal static bool TryGetLineVerdict(string text, int position, out bool verdict) {
            var entry = t_scopeDepth == 0 ? null : t_lineVerdict;
            if (entry != null && entry.Text != null && ReferenceEquals(entry.Text, text) && position >= entry.LineStart && position <= entry.LineEnd) {
                verdict = entry.Verdict;
                return true;
            }

            verdict = false;
            return false;
        }

        internal static void SetLineVerdict(string text, int lineStart, int lineEnd, bool verdict) {
            if (t_scopeDepth != 0) {
                var entry = t_lineVerdict ??= new LineVerdict();
                entry.Text = text;
                entry.LineStart = lineStart;
                entry.LineEnd = lineEnd;
                entry.Verdict = verdict;
            }
        }

        private static void Clear<T>(T?[]? entries) where T : class {
            if (entries != null) {
                Array.Clear(entries, 0, entries.Length);
            }
        }

        private static void MoveToFront<T>(T?[] entries, int index) where T : class {
            T? entry = entries[index];
            for (int i = index; i > 0; i--) {
                entries[i] = entries[i - 1];
            }

            entries[0] = entry;
        }

        private static void InsertAtFront<T>(T?[] entries, T entry) where T : class {
            for (int i = entries.Length - 1; i > 0; i--) {
                entries[i] = entries[i - 1];
            }

            entries[0] = entry;
        }
    }

    // Matching brackets for every '[' that a scan from the start of the text reaches unescaped and
    // outside skipped spans. From such a position, a scan started there visits exactly the same
    // characters, so the stack match equals what ScanForMatchingBracket would return.
    private sealed class BracketMatchIndex {
        private static readonly char[] BracketScanStops = { '\\', '`', '<', '[', ']' };

        private readonly bool _inlineHtml;
        private readonly InlineHtmlWrapperMatchIndex? _inlineHtmlWrapperMatches;
        private readonly int[] _openPositions;
        private readonly int[] _closePositions;

        private BracketMatchIndex(string text, bool inlineHtml, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches, int[] openPositions, int[] closePositions) {
            Text = text;
            _inlineHtml = inlineHtml;
            _inlineHtmlWrapperMatches = inlineHtmlWrapperMatches;
            _openPositions = openPositions;
            _closePositions = closePositions;
        }

        internal string Text { get; }

        internal bool Matches(string text, bool inlineHtml, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches) =>
            ReferenceEquals(Text, text)
            && _inlineHtml == inlineHtml
            && ReferenceEquals(_inlineHtmlWrapperMatches, inlineHtmlWrapperMatches);

        internal bool TryGetMatch(int openIndex, out int closeIndex) {
            int slot = Array.BinarySearch(_openPositions, openIndex);
            if (slot < 0) {
                closeIndex = -1;
                return false;
            }

            closeIndex = _closePositions[slot];
            return true;
        }

        internal static BracketMatchIndex Build(
            string text,
            MarkdownLoadOptions? options,
            bool inlineHtml,
            InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches) {
            // Scratch lists are reused per thread; a nested build (none is expected) takes fresh ones.
            var openPositions = t_openPositions ?? new List<int>();
            var closePositions = t_closePositions ?? new List<int>();
            var openSlots = t_openSlots ?? new Stack<int>();
            t_openPositions = null;
            t_closePositions = null;
            t_openSlots = null;
            openPositions.Clear();
            closePositions.Clear();
            openSlots.Clear();
            var effectiveInlineHtmlWrapperMatches = inlineHtmlWrapperMatches;

            // Only escapes, brackets and the starts of skipped spans ('`' and '<') change the scan.
            int i = 0;
            while (i < text.Length) {
                i = text.IndexOfAny(BracketScanStops, i);
                if (i < 0) {
                    break;
                }

                char c = text[i];
                if (c == '\\') {
                    i += 2;
                    continue;
                }

                if (c == '<' && effectiveInlineHtmlWrapperMatches == null && inlineHtml) {
                    effectiveInlineHtmlWrapperMatches = BuildInlineHtmlWrapperMatchIndex(text);
                }

                if (TrySkipLinkLabelInlineSpan(text, i, options, out int spanConsumed, effectiveInlineHtmlWrapperMatches)) {
                    i += spanConsumed;
                    continue;
                }

                if (c == '[') {
                    openSlots.Push(openPositions.Count);
                    openPositions.Add(i);
                    closePositions.Add(-1);
                } else if (c == ']' && openSlots.Count > 0) {
                    closePositions[openSlots.Pop()] = i;
                }

                i++;
            }

            var index = new BracketMatchIndex(text, inlineHtml, inlineHtmlWrapperMatches, openPositions.ToArray(), closePositions.ToArray());
            if (openPositions.Capacity <= MaximumRetainedScratchCapacity) {
                t_openPositions = openPositions;
                t_closePositions = closePositions;
                t_openSlots = openSlots;
            }

            return index;
        }

        private const int MaximumRetainedScratchCapacity = 1024;
        [ThreadStatic] private static List<int>? t_openPositions;
        [ThreadStatic] private static List<int>? t_closePositions;
        [ThreadStatic] private static Stack<int>? t_openSlots;
    }

    // Positions of maximal backtick runs ordered by run length and then by position, packed as
    // (length << 32) | start so one binary search finds the first run of a length after a position.
    private sealed class BacktickRunIndex {
        private readonly long[] _runs;

        private BacktickRunIndex(string text, long[] runs) {
            Text = text;
            _runs = runs;
        }

        internal string Text { get; }

        // Returns the start of the first maximal run of exactly <paramref name="length"/> backticks
        // starting at or after <paramref name="from"/>, which must not be inside a run.
        internal int FindRun(int from, int length) {
            long key = ((long)length << 32) | (uint)Math.Max(0, from);
            int slot = Array.BinarySearch(_runs, key);
            if (slot < 0) {
                slot = ~slot;
            }

            if (slot >= _runs.Length || (int)(_runs[slot] >> 32) != length) {
                return -1;
            }

            return (int)(_runs[slot] & 0xFFFFFFFF);
        }

        internal static BacktickRunIndex Build(string text) {
            int count = 0;
            for (int i = text.IndexOf('`'); i >= 0 && i < text.Length; i = text.IndexOf('`', i)) {
                while (i < text.Length && text[i] == '`') {
                    i++;
                }

                count++;
                if (i >= text.Length) {
                    break;
                }
            }

            var runs = new long[count];
            int index = 0;
            for (int i = text.IndexOf('`'); i >= 0 && i < text.Length; i = text.IndexOf('`', i)) {
                int start = i;
                while (i < text.Length && text[i] == '`') {
                    i++;
                }

                runs[index++] = ((long)(i - start) << 32) | (uint)start;
                if (i >= text.Length) {
                    break;
                }
            }

            Array.Sort(runs);
            return new BacktickRunIndex(text, runs);
        }
    }
}