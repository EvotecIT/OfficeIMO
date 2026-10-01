using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IElement? _activeColumnNoteOwner;

    private IReadOnlyList<HtmlRenderFlowBlock> BuildColumnChildrenWithNotes(
        IElement element, double width, HtmlRenderBoxStyle style, int depth) {
        IElement? previous = _activeColumnNoteOwner;
        _activeColumnNoteOwner = _options.Mode == HtmlRenderMode.Paged ? element : null;
        try {
            return BuildMultiColumnChildBlocks(element, element.ChildNodes, width, style, depth);
        } finally {
            _activeColumnNoteOwner = previous;
        }
    }

    private IElement? ResolveColumnNoteOwner(IElement element) {
        if (_activeColumnNoteOwner == null || !_computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)
            || !string.Equals(computed.GetValue("float-reference").Trim(), "column", StringComparison.OrdinalIgnoreCase)) return null;
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (ReferenceEquals(parent, _activeColumnNoteOwner)) return parent;
            if (_computedStyles.Elements.TryGetValue(parent, out HtmlComputedStyle? parentStyle)) {
                string count = parentStyle.GetValue("column-count").Trim();
                string width = parentStyle.GetValue("column-width").Trim();
                if ((count.Length > 0 && count != "auto") || (width.Length > 0 && width != "auto")) return null;
            }
        }
        return null;
    }

    /// <summary>
    /// Reserves each note in its call's column. A call deferred by reservation
    /// cannot move back during reflow; its preceding legal body fragment remains.
    /// </summary>
    private MultiColumnPlan ResolveColumnNoteLayout(
        IElement owner, IReadOnlyList<HtmlRenderFlowBlock> children, MultiColumnPlan body,
        double width, double height) {
        if (_footnoteEntries.Count == 0) return body;
        HtmlFootnoteEntry[] entries = _footnoteEntries.Values.Where(e => ReferenceEquals(e.ColumnOwner, owner))
            .OrderBy(e => e.Number).ToArray();
        if (entries.Length == 0) return body;
        ColumnNotePlan notes = PlanColumnNotes(body, entries, width, height, null);
        int maximumPasses = Math.Min(64, Math.Max(8, entries.Length + 4));
        for (int pass = 0; pass < maximumPasses; pass++) {
            CheckCancellation();
            body = BuildMultiColumnPlan(children, height, _options.MaxColumnCount, throwOnLimit: true, notes);
            ColumnNotePlan next = PlanColumnNotes(body, entries, width, height, notes);
            if (next.EquivalentTo(notes)) return AppendColumnNoteFragments(body, next, width, height);
            notes = next;
        }
        throw new HtmlDomLimitException(HtmlRenderDiagnosticCodes.PaginationConvergenceLimitExceeded,
            "Column-note reflow exceeded its bounded convergence limit.", "ColumnNoteReflowPasses", maximumPasses + 1, maximumPasses);
    }

    private ColumnNotePlan PlanColumnNotes(
        MultiColumnPlan body, IReadOnlyList<HtmlFootnoteEntry> entries, double width, double height,
        ColumnNotePlan? previous) {
        var chunks = new List<ColumnNoteChunk>();
        var reserved = new Dictionary<int, double>();
        var minimumCalls = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (HtmlFootnoteEntry entry in entries) {
            CheckCancellation();
            string callName = FootnoteCallDestination(entry.Number);
            int column = FindColumnNoteCall(body, callName, height, out double callHeight);
            if (column < 0) continue;
            if (previous != null && previous.MinimumCalls.TryGetValue(callName, out int minimum)) column = Math.Max(column, minimum);
            minimumCalls[callName] = column;
            int callColumn = column;
            RelayoutFootnoteEntry(entry, width);
            double offset = 0D;
            while (offset < entry.Block.Height - 0.0001D) {
                CheckCancellation();
                ChargeLayoutOperation("column note reservation");
                EnsureMultiColumnLimit(column + 1);
                reserved.TryGetValue(column, out double occupied);
                bool first = !reserved.ContainsKey(column);
                double gap = first ? FootnoteSeparatorGap : 2D;
                double minimumBody = column == callColumn
                    ? Math.Min(height, Math.Max(callHeight,
                        Math.Min(height * 0.5D, Math.Max(12D, _options.DefaultFontSize * 1.5D)))) : 0D;
                double capacity = Math.Max(0D, height - minimumBody - occupied - gap);
                if (capacity <= 0.01D) { column++; continue; }
                double end = FindFragmentEnd(entry.Block, offset, capacity, fullPageHeight: height);
                if (end <= offset + 0.0001D) {
                    if (column == callColumn || occupied > 0.0001D) { column++; continue; }
                    // An atomic note fragment larger than a fresh column remains
                    // whole; the containing column set carries its paint extent.
                    end = FindNextColumnBreak(entry.Block, offset);
                }
                double chunkHeight = Math.Max(end - offset, ResolveFootnoteMarkerHeight(entry));
                chunks.Add(new ColumnNoteChunk(entry, column, offset, end, chunkHeight, first));
                reserved[column] = occupied + gap + chunkHeight;
                offset = end;
                if (offset < entry.Block.Height - 0.0001D) column++;
            }
        }
        return new ColumnNotePlan(chunks, reserved, minimumCalls);
    }

    private int FindColumnNoteCall(MultiColumnPlan body, string callName, double height, out double callHeight) {
        callHeight = 0D;
        foreach (MultiColumnFragment fragment in body.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("column note anchor lookup");
            HtmlRenderNamedDestination? call = EnumeratePageFloatVisuals(SliceBlockVisuals(fragment.Block, fragment.Start, fragment.End))
                .OfType<HtmlRenderNamedDestination>().FirstOrDefault(d => d.Name == callName);
            if (call == null) continue;
            double anchor = fragment.Start + call.LayoutY;
            double before = FindFragmentEnd(fragment.Block, 0D, Math.Max(0D, anchor), anchor, fullPageHeight: height);
            double after = FindNextColumnBreak(fragment.Block, anchor);
            callHeight = Math.Max(0D, after - before);
            if (fragment.Block.AvoidBreakInside) callHeight = Math.Max(callHeight, fragment.Block.Height);
            foreach (HtmlRenderAvoidBreakRange range in fragment.Block.AvoidBreakRanges) {
                if (anchor >= range.Start && anchor < range.End && range.End - range.Start <= height)
                    callHeight = Math.Max(callHeight, range.End - range.Start);
            }
            return fragment.Column;
        }
        return -1;
    }

    private double RestrictFragmentBeforeDeferredColumnCall(
        HtmlRenderFlowBlock child, double start, double end, int column, double height, ColumnNotePlan notes) {
        foreach (HtmlRenderNamedDestination call in EnumeratePageFloatVisuals(child.Visuals).OfType<HtmlRenderNamedDestination>()) {
            CheckCancellation();
            ChargeLayoutOperation("deferred column note anchor");
            if (!notes.MinimumCalls.TryGetValue(call.Name, out int minimum) || minimum <= column
                || call.LayoutY < start - 0.0001D || call.LayoutY >= end - 0.0001D) continue;
            end = FindFragmentEnd(child, start, Math.Max(0D, call.LayoutY - start),
                Math.Min(end, call.LayoutY), fullPageHeight: height);
        }
        return end;
    }

    /// <summary>
    /// A column set with reserved notes is one physical fragmentainer. Keep its
    /// call and note together when that set fits a fresh page; overflow sets
    /// retain their authored row-boundary page breaks.
    /// </summary>
    private IEnumerable<HtmlRenderAvoidBreakRange> ResolveColumnNoteKeepRanges(
        IElement owner, IReadOnlyList<double> rowBreaks, double contentY, double contentHeight) {
        if (!_footnoteEntries.Values.Any(entry => ReferenceEquals(entry.ColumnOwner, owner))) yield break;
        double start = 0D;
        foreach (double end in rowBreaks.Concat(new[] { contentHeight })) {
            yield return new HtmlRenderAvoidBreakRange(contentY + start, contentY + end);
            start = end;
        }
    }

    private MultiColumnPlan AppendColumnNoteFragments(MultiColumnPlan body, ColumnNotePlan notes, double width, double height) {
        var fragments = new List<MultiColumnFragment>(body.Fragments);
        double extent = Math.Max(height, notes.Reservations.Values.DefaultIfEmpty(0D).Max());
        var cursors = notes.Reservations.ToDictionary(pair => pair.Key, pair => extent - pair.Value);
        foreach (ColumnNoteChunk note in notes.Chunks) {
            CheckCancellation();
            ChargeLayoutOperation("column note painting");
            double gap = note.First ? FootnoteSeparatorGap : 2D;
            double cursor = cursors[note.Column];
            var visuals = new List<HtmlRenderVisual>();
            if (note.First) {
                double separatorWidth = Math.Min(width, Math.Max(24D, width * 0.25D));
                OfficeShape shape = OfficeShape.Line(0D, 2D, separatorWidth, 2D);
                shape.Height = 0.75D; shape.FillColor = null;
                shape.StrokeColor = note.Entry.Style.Color; shape.StrokeWidth = 0.75D;
                visuals.Add(new HtmlRenderSemanticGroup(HtmlRenderSemanticGroupRole.Artifact, 0D, 2D, separatorWidth, 0.75D,
                    new HtmlRenderVisual[] { new HtmlRenderShape(shape, 0D, 2D, _paintOrder++, source: "footnote-separator") },
                    _paintOrder++, "footnote-separator"));
            }
            var chunk = new HtmlFootnoteChunk(note.Entry.Element, 0, note.Start, note.End, note.First, note.Height);
            AddFootnoteChunkVisuals(visuals, note.Entry, chunk, 0D, gap, width);
            double blockHeight = gap + note.Height;
            double[] breaks = note.Entry.Block.BreakOffsets.Where(cut => cut > note.Start && cut <= note.End)
                .Select(cut => gap + cut - note.Start).Concat(new[] { blockHeight }).Distinct().OrderBy(cut => cut).ToArray();
            var block = new HtmlRenderFlowBlock(width, blockHeight, visuals, HtmlPageBreakTarget.None,
                HtmlPageBreakTarget.None, false, HtmlRenderStyleResolver.DescribeSource(note.Entry.Element) + ":column-note", breaks);
            fragments.Add(new MultiColumnFragment(block, 0D, blockHeight, note.Column, cursor));
            cursors[note.Column] += blockHeight;
        }
        int count = Math.Max(body.ColumnCount, notes.Reservations.Keys.DefaultIfEmpty(-1).Max() + 1);
        return new MultiColumnPlan(fragments, count, Math.Max(body.UsedHeight, extent));
    }

    private sealed record ColumnNoteChunk(HtmlFootnoteEntry Entry, int Column, double Start, double End, double Height, bool First);

    private sealed class ColumnNotePlan {
        internal ColumnNotePlan(IReadOnlyList<ColumnNoteChunk> chunks, IReadOnlyDictionary<int, double> reservations,
            IReadOnlyDictionary<string, int> minimumCalls) {
            Chunks = chunks; Reservations = reservations; MinimumCalls = minimumCalls;
        }
        internal IReadOnlyList<ColumnNoteChunk> Chunks { get; }
        internal IReadOnlyDictionary<int, double> Reservations { get; }
        internal IReadOnlyDictionary<string, int> MinimumCalls { get; }
        internal double Reserved(int column) => Reservations.TryGetValue(column, out double height) ? height : 0D;
        internal bool EquivalentTo(ColumnNotePlan other) =>
            MinimumCalls.Count == other.MinimumCalls.Count && MinimumCalls.All(p => other.MinimumCalls.TryGetValue(p.Key, out int n) && p.Value == n)
            && Reservations.Count == other.Reservations.Count && Reservations.All(p => Math.Abs(p.Value - other.Reserved(p.Key)) < 0.0001D)
            && Chunks.SequenceEqual(other.Chunks);
    }
}
