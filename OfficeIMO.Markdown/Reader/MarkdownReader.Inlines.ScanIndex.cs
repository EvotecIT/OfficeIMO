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
                t_scopeDepth = 0;
                t_bracketIndexes = null;
                t_backtickIndexes = null;
                t_forwardSearches = null;
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
            var openPositions = new List<int>();
            var closePositions = new List<int>();
            var openSlots = new Stack<int>();
            var effectiveInlineHtmlWrapperMatches = inlineHtmlWrapperMatches;
            bool escaped = false;

            for (int i = 0; i < text.Length; i++) {
                char c = text[i];
                if (escaped) {
                    escaped = false;
                    continue;
                }

                if (c == '\\') {
                    escaped = true;
                    continue;
                }

                if (c == '<' && effectiveInlineHtmlWrapperMatches == null && inlineHtml) {
                    effectiveInlineHtmlWrapperMatches = BuildInlineHtmlWrapperMatchIndex(text);
                }

                if (TrySkipLinkLabelInlineSpan(text, i, options, out int spanConsumed, effectiveInlineHtmlWrapperMatches)) {
                    i += spanConsumed - 1;
                    continue;
                }

                if (c == '[') {
                    openSlots.Push(openPositions.Count);
                    openPositions.Add(i);
                    closePositions.Add(-1);
                    continue;
                }

                if (c == ']' && openSlots.Count > 0) {
                    closePositions[openSlots.Pop()] = i;
                }
            }

            return new BracketMatchIndex(text, inlineHtml, inlineHtmlWrapperMatches, openPositions.ToArray(), closePositions.ToArray());
        }
    }

    // Positions of maximal backtick runs grouped by run length, in text order.
    private sealed class BacktickRunIndex {
        private readonly Dictionary<int, List<int>> _runsByLength;

        private BacktickRunIndex(string text, Dictionary<int, List<int>> runsByLength) {
            Text = text;
            _runsByLength = runsByLength;
        }

        internal string Text { get; }

        // Returns the start of the first maximal run of exactly <paramref name="length"/> backticks
        // starting at or after <paramref name="from"/>, which must not be inside a run.
        internal int FindRun(int from, int length) {
            if (!_runsByLength.TryGetValue(length, out var runs)) {
                return -1;
            }

            int slot = runs.BinarySearch(from);
            if (slot < 0) {
                slot = ~slot;
            }

            return slot < runs.Count ? runs[slot] : -1;
        }

        internal static BacktickRunIndex Build(string text) {
            var runsByLength = new Dictionary<int, List<int>>();
            for (int i = text.IndexOf('`'); i >= 0 && i < text.Length; i = text.IndexOf('`', i)) {
                int start = i;
                while (i < text.Length && text[i] == '`') {
                    i++;
                }

                int length = i - start;
                if (!runsByLength.TryGetValue(length, out var runs)) {
                    runs = new List<int>();
                    runsByLength[length] = runs;
                }

                runs.Add(start);
                if (i >= text.Length) {
                    break;
                }
            }

            return new BacktickRunIndex(text, runsByLength);
        }
    }
}
