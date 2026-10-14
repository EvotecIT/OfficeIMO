using System.Text;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Maps real text adjacency across styled runs without changing their shaping,
    /// logical offsets, or inline-box ownership. Only the next word is previewed.
    /// </summary>
    private InlineWordFitting CreateInlineWordFitting(IReadOnlyList<HtmlInlineRun> runs,
        HtmlRenderBoxStyle paragraphStyle, Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector = null,
        Func<int>? consecutiveHyphenatedLines = null) =>
        new(() => BuildInlineWordFitting(runs, paragraphStyle, styleSelector), consecutiveHyphenatedLines);

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
            if (run.Text.Length == 0) {
                if (run.InlineEdgeBoundary != null) mapped.Add(new InlineWordSourceRun(run, source.Length));
                continue;
            }
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
        InlineWordSourceRun[] textRuns = mapped.Where(item => item.Run.Text.Length > 0).ToArray();
        HashSet<int>? nonCjk = textRuns.Any(item => item.Run.Style.WordBreak == "keep-all")
            ? new HashSet<int>(GetHtmlPreferredBreakPositions(text, allowCjkBreaks: false)) : null;
        int owner = 0;
        foreach (int point in preferred) {
            while (owner + 1 < textRuns.Length && textRuns[owner + 1].Start <= point) owner++;
            if (textRuns.Length == 0 || textRuns[owner].Run.Style.WordBreak != "keep-all"
                || nonCjk!.Contains(point)) breaks[point] = true;
        }
        CheckCancellation();
        return new InlineWordMap(this, mapped.ToArray(), breaks, paragraphStyle.PreventTextWrapping, styleSelector);
    }

    private sealed class InlineWordFitting {
        private readonly Lazy<InlineWordMap> _map;
        private readonly Func<int>? _consecutiveHyphenatedLines;
        internal InlineWordFitting(Func<InlineWordMap> createMap, Func<int>? consecutiveHyphenatedLines) {
            _map = new Lazy<InlineWordMap>(createMap);
            _consecutiveHyphenatedLines = consecutiveHyphenatedLines;
        }
        internal bool CanBreakBefore(HtmlInlineRun run, int sourceOffset) => _map.Value.CanBreakBefore(run, sourceOffset);
        internal bool HasBreakInside(HtmlInlineRun run, int sourceStart, int sourceEnd) => _map.Value.HasBreakInside(run, sourceStart, sourceEnd);
        internal double PreviewWithContinuation(InlineLine line, HtmlInlineRun run, int sourceEnd,
            double contentWidth, bool completesToken = true) =>
            _map.Value.PreviewWithContinuation(line, run, sourceEnd, contentWidth, completesToken,
                _consecutiveHyphenatedLines);
    }

    private sealed class InlineWordMap {
        private readonly HtmlRenderLayoutEngine _owner;
        private readonly InlineWordSourceRun[] _runs;
        private readonly Dictionary<HtmlInlineRun, int> _indices = new();
        private readonly int[] _previousBreak;
        private readonly int[] _nextBreak;
        private readonly Func<HtmlInlineRun, HtmlRenderBoxStyle>? _styleSelector;
        private readonly bool _preventWrapping;
        private readonly Dictionary<(int Run, int Offset), InlineWordPreview[]> _continuations = new();
        private readonly Dictionary<(HtmlInlineRun Run, int End), int[]> _hyphenationBreaks = new();

        internal InlineWordMap(HtmlRenderLayoutEngine owner, InlineWordSourceRun[] runs,
            bool[] breaks, bool preventWrapping, Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector) {
            _owner = owner;
            _runs = runs;
            _styleSelector = styleSelector;
            _preventWrapping = preventWrapping;
            for (int index = 0; index < runs.Length; index++) _indices.Add(runs[index].Run, index);
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
            !_indices.TryGetValue(run, out int index)
            || _previousBreak[_runs[index].Start + sourceOffset] == _runs[index].Start + sourceOffset;

        internal bool HasBreakInside(HtmlInlineRun run, int sourceStart, int sourceEnd) =>
            sourceEnd > sourceStart && _indices.TryGetValue(run, out int index)
            && _previousBreak[_runs[index].Start + sourceEnd - 1] > _runs[index].Start + sourceStart;

        /// <summary>Previews a source-mapped tail without consuming scopes or altering future runs.</summary>
        internal double PreviewWithContinuation(InlineLine line, HtmlInlineRun run, int sourceEnd,
            double contentWidth, bool completesToken, Func<int>? consecutiveHyphenatedLines) {
            if (!_indices.TryGetValue(run, out int runIndex)) return line.PreviewAdvance(run, contentWidth, completesToken);
            int offset = _runs[runIndex].Start + sourceEnd;
            int end = _nextBreak[offset];
            var key = (runIndex, offset);
            if (!_continuations.TryGetValue(key, out InlineWordPreview[]? pieces)) {
                var tail = new List<InlineWordPreview>();
                int first = sourceEnd < run.Text.Length ? runIndex : runIndex + 1;
                for (int index = first; index < _runs.Length && _runs[index].Start <= end; index++) {
                    _owner.CheckCancellation();
                    _owner.ChargeLayoutOperation("inline word continuation");
                    InlineWordSourceRun source = _runs[index];
                    if (source.Run.InlineEdgeBoundary != null) {
                        tail.Add(new InlineWordPreview(source.Run, 0D, true, Array.Empty<double>()));
                        continue;
                    }
                    int from = Math.Max(offset, source.Start) - source.Start;
                    int to = Math.Min(end, source.End) - source.Start;
                    if (to <= from) continue;
                    string value = source.Run.Text.Substring(from, to - from).Replace("\u00AD", string.Empty);
                    if (to == source.Run.Text.Length && source.Run.EndsFirstLine) value += source.Run.FirstLineHyphen;
                    HtmlRenderBoxStyle style = _styleSelector?.Invoke(source.Run) ?? source.Run.Style;
                    double width = _owner.MeasureInlineText(value, style);
                    var hyphenatedWidths = new List<double>();
                    // The current token has already attempted its own hyphenation. Only
                    // later runs will reach the normal hyphenation path after this preview.
                    if (index != runIndex) {
                        foreach (int point in GetHyphenationBreaks(source.Run, to)) {
                            if (point <= from || point >= to) continue;
                            _owner.CheckCancellation();
                            _owner.ChargeLayoutOperations(1L + (point - from + 31L) / 32L, "inline word hyphenation preview");
                            string prefix = source.Run.Text.Substring(from, point - from).Replace("\u00AD", string.Empty);
                            hyphenatedWidths.Add(_owner.MeasureInlineText(prefix + source.Run.Style.HyphenateCharacter, style));
                        }
                    }
                    tail.Add(new InlineWordPreview(source.Run, width, to == source.Run.Text.Length, hyphenatedWidths.ToArray()));
                }
                pieces = tail.ToArray();
                _continuations.Add(key, pieces);
            }
            if (pieces.Length == 0) return line.PreviewAdvance(run, contentWidth, completesToken);
            int hyphenatedLines = pieces.Any(piece => piece.HyphenatedWidths.Count > 0)
                ? consecutiveHyphenatedLines?.Invoke() ?? 0 : 0;
            return line.PreviewWordAdvance(run, contentWidth, completesToken, pieces, hyphenatedLines);
        }

        /// <summary>Uses the same prepared token and authored character limits as actual line breaking.</summary>
        private IReadOnlyList<int> GetHyphenationBreaks(HtmlInlineRun run, int sourceEnd) {
            var key = (run, sourceEnd);
            if (_hyphenationBreaks.TryGetValue(key, out int[]? cached)) return cached;
            var points = new List<int>();
            if (!_preventWrapping && !run.Style.PreventTextWrapping && run.Style.Hyphens != "none"
                && run.Style.WordBreak != "break-all") {
                int sourceOffset = 0, logicalOffset = 0;
                foreach (string token in Tokenize(run.Text, run.Style.PreserveWhitespace, run.Style.BreakSpaces)) {
                    if (sourceOffset >= sourceEnd) break;
                    _owner.CheckCancellation();
                    _owner.ChargeLayoutOperations(1L + (token.Length + 31L) / 32L, "inline word hyphenation mapping");
                    int sourceStart = sourceOffset;
                    sourceOffset += run.Style.PreserveWhitespace && token == "\n" && run.Text[sourceOffset] == '\r'
                        && sourceOffset + 1 < run.Text.Length && run.Text[sourceOffset + 1] == '\n' ? 2 : token.Length;
                    string logicalToken = SliceLogicalToken(run, token, ref logicalOffset);
                    if (IsWhitespaceToken(token)) continue;
                    HyphenationToken prepared = run.PreparedHyphenation.HasValue
                        ? run.PreparedHyphenation.Value.SliceSource(logicalOffset - token.Length, logicalToken.Length)
                        : _owner.PrepareHyphenationToken(token, logicalToken, run.Style, reportUnsupportedLanguage: false);
                    if (prepared.PaintText.Length != prepared.LogicalText.Length) continue;
                    foreach (int point in prepared.PrimaryBreaks.Concat(prepared.SecondaryBreaks)) {
                        int boundary = point < prepared.SourceBoundaries.Count ? prepared.SourceBoundaries[point] : point;
                        points.Add(sourceStart + boundary);
                    }
                }
            }
            cached = points.Distinct().OrderBy(point => point).ToArray();
            _hyphenationBreaks.Add(key, cached);
            return cached;
        }
    }

    private sealed class InlineWordSourceRun {
        internal InlineWordSourceRun(HtmlInlineRun run, int start) { Run = run; Start = start; }
        internal HtmlInlineRun Run { get; }
        internal int Start { get; }
        internal int End => Start + Run.Text.Length;
    }

    private readonly struct InlineWordPreview {
        internal InlineWordPreview(HtmlInlineRun run, double width, bool endsRun, IReadOnlyList<double> hyphenatedWidths) {
            Run = run; Width = width; EndsRun = endsRun; HyphenatedWidths = hyphenatedWidths;
        }
        internal HtmlInlineRun Run { get; }
        internal double Width { get; }
        internal bool EndsRun { get; }
        internal IReadOnlyList<double> HyphenatedWidths { get; }
    }

    private sealed partial class InlineLine {
        /// <summary>Extends the canonical advance preview; live scopes and their first/last content remain untouched.</summary>
        internal double PreviewWordAdvance(HtmlInlineRun run, double contentWidth, bool completesToken,
            IReadOnlyList<InlineWordPreview> continuation, int consecutiveHyphenatedLines) {
            var preview = new InlineAdvancePreview(this);
            preview.AddContent(run, contentWidth);
            double hyphenated = double.PositiveInfinity;
            HtmlInlineRun lastRun = run;
            bool completesLastRun = completesToken && run.InlineTokenEndsRun;
            foreach (InlineWordPreview piece in continuation) {
                if (piece.Run.InlineEdgeBoundary is { } boundary) {
                    if (!boundary.Opening) preview.CloseScope(piece.Run);
                } else {
                    if (!piece.Run.Style.HyphenateLimitLines.HasValue
                        || consecutiveHyphenatedLines < piece.Run.Style.HyphenateLimitLines.Value) {
                        foreach (double width in piece.HyphenatedWidths)
                            hyphenated = Math.Min(hyphenated, preview.PreviewContent(piece.Run, width));
                    }
                    preview.AddContent(piece.Run, piece.Width);
                }
                lastRun = piece.Run;
                completesLastRun = piece.EndsRun;
            }
            return Math.Min(hyphenated, preview.Width + (completesLastRun ? lastRun.InlineClosingAdvance : 0D));
        }
    }
}
