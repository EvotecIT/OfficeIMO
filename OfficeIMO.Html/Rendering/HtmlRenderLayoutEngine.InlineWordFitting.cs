using System.Text;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Maps real text adjacency across styled runs without changing their shaping,
    /// logical offsets, or inline-box ownership. Only the next word is previewed.
    /// </summary>
    private InlineWordFitting CreateInlineWordFitting(IReadOnlyList<HtmlInlineRun> runs,
        HtmlRenderBoxStyle paragraphStyle, Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector = null) =>
        new(() => BuildInlineWordFitting(runs, paragraphStyle, styleSelector));

    private InlineWordMap BuildInlineWordFitting(IReadOnlyList<HtmlInlineRun> runs,
        HtmlRenderBoxStyle paragraphStyle, Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector = null) {
        var source = new StringBuilder();
        var mapped = new List<InlineWordSourceRun>();
        var explicitBoundaries = new List<int> { 0 };
        var wordBreakElements = new HashSet<object>();
        foreach (HtmlInlineRun run in runs) {
            CheckCancellation();
            ChargeLayoutOperations(1L + (run.Text.Length + 31L) / 32L, "inline word source mapping");
            bool wordBreak = run.OwnerElement?.LocalName == "wbr"
                && !paragraphStyle.PreventTextWrapping && !run.Style.PreventTextWrapping
                && wordBreakElements.Add(run.OwnerElement);
            bool suppressedBreak = run.IsSuppressedWhitespaceBreak
                && !paragraphStyle.PreventTextWrapping && !run.Style.PreventTextWrapping;
            if (wordBreak || suppressedBreak || run.AtomicBlock != null && !run.IsFlowMarker || run.IsBlockInterruption) {
                explicitBoundaries.Add(source.Length);
                source.Append('\u2028');
                explicitBoundaries.Add(source.Length);
            }
            if (run.Text.Length == 0) continue;
            mapped.Add(new InlineWordSourceRun(run, source.Length));
            for (int offset = 0; offset < run.Text.Length; offset++) {
                if ((offset & 1023) == 0) CheckCancellation();
                char value = run.Text[offset];
                if (IsCollapsibleCssWhitespace(value) || value == '\u2028') {
                    explicitBoundaries.Add(source.Length);
                    explicitBoundaries.Add(source.Length + 1);
                }
                source.Append(value);
            }
            if (run.EndsFirstLine) {
                explicitBoundaries.Add(source.Length);
                source.Append('\u2028');
                explicitBoundaries.Add(source.Length);
            }
        }

        string text = source.ToString();
        var breaks = new bool[text.Length + 1];
        foreach (int point in explicitBoundaries) breaks[point] = true;
        breaks[text.Length] = true;
        IReadOnlyList<int> preferred = GetHtmlPreferredBreakPositions(text, allowCjkBreaks: true);
        HashSet<int>? nonCjk = mapped.Any(item => item.Run.Style.WordBreak == "keep-all")
            ? new HashSet<int>(GetHtmlPreferredBreakPositions(text, allowCjkBreaks: false)) : null;
        int owner = 0;
        foreach (int point in preferred) {
            while (owner + 1 < mapped.Count && mapped[owner + 1].Start <= point) owner++;
            if (mapped.Count == 0 || mapped[owner].Run.Style.WordBreak != "keep-all"
                || nonCjk!.Contains(point)) breaks[point] = true;
        }
        CheckCancellation();
        return new InlineWordMap(this, mapped.ToArray(), breaks, styleSelector);
    }

    private sealed class InlineWordFitting {
        private readonly Lazy<InlineWordMap> _map;
        internal InlineWordFitting(Func<InlineWordMap> createMap) { _map = new Lazy<InlineWordMap>(createMap); }
        internal bool CanBreakBefore(HtmlInlineRun run, int sourceOffset) => _map.Value.CanBreakBefore(run, sourceOffset);
        internal bool HasBreakInside(HtmlInlineRun run, int sourceStart, int sourceEnd) => _map.Value.HasBreakInside(run, sourceStart, sourceEnd);
        internal double PreviewWithContinuation(InlineLine line, HtmlInlineRun run, int sourceEnd,
            double contentWidth, bool completesToken = true) =>
            _map.Value.PreviewWithContinuation(line, run, sourceEnd, contentWidth, completesToken);
    }

    private sealed class InlineWordMap {
        private readonly HtmlRenderLayoutEngine _owner;
        private readonly InlineWordSourceRun[] _runs;
        private readonly Dictionary<HtmlInlineRun, int> _starts = new();
        private readonly int[] _previousBreak;
        private readonly int[] _nextBreak;
        private readonly Func<HtmlInlineRun, HtmlRenderBoxStyle>? _styleSelector;
        private readonly Dictionary<int, InlineWordPreview[]> _continuations = new();

        internal InlineWordMap(HtmlRenderLayoutEngine owner, InlineWordSourceRun[] runs,
            bool[] breaks, Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector) {
            _owner = owner;
            _runs = runs;
            _styleSelector = styleSelector;
            foreach (InlineWordSourceRun run in runs) _starts.Add(run.Run, run.Start);
            _previousBreak = new int[breaks.Length];
            _nextBreak = new int[breaks.Length];
            int previous = 0, next = breaks.Length - 1;
            for (int index = 0; index < breaks.Length; index++) {
                if ((index & 1023) == 0) owner.CheckCancellation();
                if (breaks[index]) previous = index;
                _previousBreak[index] = previous;
                int reverse = breaks.Length - index - 1;
                if (breaks[reverse]) next = reverse;
                _nextBreak[reverse] = next;
            }
        }

        internal bool CanBreakBefore(HtmlInlineRun run, int sourceOffset) =>
            !_starts.TryGetValue(run, out int start) || _previousBreak[start + sourceOffset] == start + sourceOffset;

        internal bool HasBreakInside(HtmlInlineRun run, int sourceStart, int sourceEnd) =>
            sourceEnd > sourceStart && _starts.TryGetValue(run, out int start)
            && _previousBreak[start + sourceEnd - 1] > start + sourceStart;

        /// <summary>Previews a source-mapped tail without consuming scopes or altering future runs.</summary>
        internal double PreviewWithContinuation(InlineLine line, HtmlInlineRun run, int sourceEnd,
            double contentWidth, bool completesToken = true) {
            if (!_starts.TryGetValue(run, out int start)) return line.PreviewAdvance(run, contentWidth, completesToken);
            int offset = start + sourceEnd;
            int end = _nextBreak[offset];
            if (end == offset) return line.PreviewAdvance(run, contentWidth, completesToken);
            if (!_continuations.TryGetValue(offset, out InlineWordPreview[]? pieces)) {
                var tail = new List<InlineWordPreview>();
                int low = 0, high = _runs.Length;
                while (low < high) {
                    int middle = low + (high - low) / 2;
                    if (_runs[middle].End <= offset) low = middle + 1;
                    else high = middle;
                }
                for (int index = low; index < _runs.Length && _runs[index].Start < end; index++) {
                    _owner.CheckCancellation();
                    _owner.ChargeLayoutOperation("inline word continuation");
                    InlineWordSourceRun source = _runs[index];
                    int from = Math.Max(offset, source.Start) - source.Start;
                    int to = Math.Min(end, source.End) - source.Start;
                    if (to <= from) continue;
                    string value = source.Run.Text.Substring(from, to - from).Replace("\u00AD", string.Empty);
                    double width = _owner.MeasureInlineText(value, _styleSelector?.Invoke(source.Run) ?? source.Run.Style);
                    tail.Add(new InlineWordPreview(source.Run, width, to == source.Run.Text.Length));
                }
                pieces = tail.ToArray();
                _continuations.Add(offset, pieces);
            }
            return line.PreviewWordAdvance(run, contentWidth, completesToken, pieces);
        }
    }

    private sealed class InlineWordSourceRun {
        internal InlineWordSourceRun(HtmlInlineRun run, int start) { Run = run; Start = start; }
        internal HtmlInlineRun Run { get; }
        internal int Start { get; }
        internal int End => Start + Run.Text.Length;
    }

    private readonly struct InlineWordPreview {
        internal InlineWordPreview(HtmlInlineRun run, double width, bool endsRun) {
            Run = run; Width = width; EndsRun = endsRun;
        }
        internal HtmlInlineRun Run { get; }
        internal double Width { get; }
        internal bool EndsRun { get; }
    }

    private sealed partial class InlineLine {
        /// <summary>Extends the canonical advance preview; live scopes and their first/last content remain untouched.</summary>
        internal double PreviewWordAdvance(HtmlInlineRun run, double contentWidth, bool completesToken,
            IReadOnlyList<InlineWordPreview> continuation) {
            double result = PreviewAdvance(run, contentWidth, completesToken);
            var present = new HashSet<HtmlInlineEdgeScope>(_inlineEdges.Keys);
            foreach (HtmlInlineEdgeScope scope in run.InlineEdgeScopes) present.Add(scope);
            foreach (InlineWordPreview piece in continuation) {
                result += piece.Width;
                foreach (HtmlInlineEdgeScope scope in piece.Run.InlineEdgeScopes) {
                    if (!present.Add(scope)) continue;
                    if (scope.Clone || scope.FirstContent == null) result += scope.Start;
                    if (scope.Clone) result += scope.End;
                }
                if (piece.EndsRun) result += piece.Run.InlineClosingAdvance;
            }
            return result;
        }
    }
}
