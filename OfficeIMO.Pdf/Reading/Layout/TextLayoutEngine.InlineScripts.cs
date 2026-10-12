namespace OfficeIMO.Pdf;

internal static partial class TextLayoutEngine {
    private sealed class InlineScriptRun {
        internal readonly int GroupIndex;
        internal readonly List<PdfTextSpan> Spans;
        internal readonly double Start;
        internal readonly double End;
        internal readonly double NormalEm;
        internal readonly double MinimumNormal;
        internal readonly double MaximumNormal;

        internal InlineScriptRun(int groupIndex, List<PdfTextSpan> spans, PdfTextBaselineCoordinates coordinates,
            Action<long>? consumeWork, Action? cancellationCheck) {
            GroupIndex = groupIndex;
            Spans = spans;
            Start = coordinates.Along(spans[0]);
            End = Start;
            MinimumNormal = double.MaxValue;
            MaximumNormal = double.MinValue;
            foreach (PdfTextSpan span in spans) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1L);
                End = Math.Max(End, coordinates.Along(span) + Math.Max(0D, span.Advance));
                TryGetInlineScriptEm(span, out _, out double normalEm);
                NormalEm = Math.Max(NormalEm, normalEm);
                double normal = coordinates.Normal(span);
                MinimumNormal = Math.Min(MinimumNormal, normal);
                MaximumNormal = Math.Max(MaximumNormal, normal);
            }
        }
    }

    private static void AttachInlineScriptRuns(List<BaselineSpanGroup> groups,
        PdfTextBaselineCoordinates coordinates, Options options, Action<long>? consumeWork,
        Action? cancellationCheck) {
        // Explicitly tighter limits keep their baseline-only meaning. The default path
        // separately recognizes adjacent smaller text without widening ordinary clustering.
        if (groups.Count <= 1 || !(options.LineMergeToleranceEm >= 0.6D && options.LineMergeMaxPoints >= 2.5D)) return;
        double maximumNormalEm = 0D, minimumNormalEm = double.MaxValue;
        int spanCount = 0;
        foreach (BaselineSpanGroup group in groups) {
            foreach (PdfTextSpan span in group.Spans) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1L);
                if (!TryGetInlineScriptEm(span, out _, out double normalEm)) return;
                spanCount++;
                maximumNormalEm = Math.Max(maximumNormalEm, normalEm);
                minimumNormalEm = Math.Min(minimumNormalEm, normalEm);
            }
        }
        if (minimumNormalEm > maximumNormalEm * 0.9D) return;
        bool explicitDirection = options.ReadingDirection != PdfReadingDirection.Auto;
        PdfReadingDirection direction = options.ReadingDirection;
        if (!explicitDirection) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(2L * spanCount * Math.Max(1,
                (int)Math.Ceiling(Math.Log(Math.Max(2, spanCount), 2D))));
            direction = PdfTextDirectionAnalysis.Resolve(PdfReadingDirection.Auto,
                groups.SelectMany(static group => group.Spans)
                    .OrderBy(span => {
                        cancellationCheck?.Invoke();
                        return span.ContentOrderKey;
                    }).ThenBy(span => {
                        cancellationCheck?.Invoke();
                        return span.PaintOrder;
                    }).Select(span => {
                        cancellationCheck?.Invoke();
                        return span.Text;
                    }));
        }

        var orderedSpans = new List<PdfTextSpan>[groups.Count];
        var runs = new List<InlineScriptRun>();
        for (int index = 0; index < groups.Count; index++) {
            cancellationCheck?.Invoke();
            int count = groups[index].Spans.Count;
            consumeWork?.Invoke((long)count * Math.Max(1, (int)Math.Ceiling(Math.Log(Math.Max(2, count), 2D))));
            orderedSpans[index] = groups[index].Spans.OrderBy(span => {
                cancellationCheck?.Invoke();
                return coordinates.Along(span);
            }).ToList();
            AddInlineScriptRuns(runs, index, orderedSpans[index], coordinates, direction, explicitDirection, consumeWork, cancellationCheck);
        }

        var attachments = new List<(InlineScriptRun Run, int GroupIndex, PdfTextSpan Anchor)>();
        foreach (InlineScriptRun run in runs) {
            cancellationCheck?.Invoke();
            if (run.NormalEm > maximumNormalEm * 0.9D) continue;
            int first = FirstBaselineAtOrBelow(groups, run.MaximumNormal + maximumNormalEm * 0.5D,
                consumeWork, cancellationCheck);
            int selected = -1;
            PdfTextSpan? selectedAnchor = null;
            for (int index = first; index < groups.Count &&
                groups[index].Normal >= run.MinimumNormal - maximumNormalEm * 0.5D; index++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1L);
                if (index == run.GroupIndex) continue;
                PdfTextSpan? anchor = FindInlineScriptAnchor(run, groups[index], orderedSpans[index],
                    coordinates, direction, explicitDirection, consumeWork, cancellationCheck);
                if (anchor is null) continue;
                if (selected >= 0) { selected = -1; break; }
                selected = index;
                selectedAnchor = anchor;
            }
            if (selected >= 0 && selectedAnchor is not null) attachments.Add((run, selected, selectedAnchor));
        }

        // Do not infer a nested chain through text that is itself being classified as a script.
        var moving = new HashSet<PdfTextSpan>(attachments.SelectMany(static item => item.Run.Spans));
        var removed = new Dictionary<int, HashSet<PdfTextSpan>>();
        foreach (var attachment in attachments) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(attachment.Run.Spans.Count);
            if (moving.Contains(attachment.Anchor)) continue;
            BaselineSpanGroup target = groups[attachment.GroupIndex];
            target.OriginalSpans ??= new List<PdfTextSpan>(target.Spans);
            target.Spans.AddRange(attachment.Run.Spans);
            if (!removed.TryGetValue(attachment.Run.GroupIndex, out HashSet<PdfTextSpan>? sourceRemoved)) {
                sourceRemoved = new HashSet<PdfTextSpan>();
                removed.Add(attachment.Run.GroupIndex, sourceRemoved);
            }
            sourceRemoved.UnionWith(attachment.Run.Spans);
        }
        foreach (var source in removed) {
            cancellationCheck?.Invoke();
            bool Remove(PdfTextSpan span) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1L);
                return source.Value.Contains(span);
            }
            groups[source.Key].Spans.RemoveAll(Remove);
            groups[source.Key].OriginalSpans?.RemoveAll(Remove);
        }
    }

    private static void AddInlineScriptRuns(List<InlineScriptRun> runs, int groupIndex,
        List<PdfTextSpan> spans, PdfTextBaselineCoordinates coordinates, PdfReadingDirection direction,
        bool explicitDirection, Action<long>? consumeWork, Action? cancellationCheck) {
        var current = new List<PdfTextSpan> { spans[0] };
        for (int index = 1; index < spans.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1L);
            PdfTextSpan previous = spans[index - 1], span = spans[index];
            double gap = coordinates.Along(span) - coordinates.Along(previous) - Math.Max(0D, previous.Advance);
            TryGetInlineScriptEm(previous, out double previousAlongEm, out _);
            TryGetInlineScriptEm(span, out double alongEm, out _);
            double threshold = Math.Max(0.5D, Math.Min(previousAlongEm, alongEm) * 0.25D);
            if (gap > threshold || InlineScriptBoundaryHasBreak(previous, span, direction, explicitDirection)) {
                runs.Add(new InlineScriptRun(groupIndex, current, coordinates, consumeWork, cancellationCheck));
                current = new List<PdfTextSpan>();
            }
            current.Add(span);
        }
        runs.Add(new InlineScriptRun(groupIndex, current, coordinates, consumeWork, cancellationCheck));
    }

    private static int FirstBaselineAtOrBelow(List<BaselineSpanGroup> groups, double normal,
        Action<long>? consumeWork, Action? cancellationCheck) {
        int low = 0, high = groups.Count;
        while (low < high) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1L);
            int middle = low + (high - low) / 2;
            if (groups[middle].Normal > normal) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    private static PdfTextSpan? FindInlineScriptAnchor(InlineScriptRun run, BaselineSpanGroup group,
        List<PdfTextSpan> spans, PdfTextBaselineCoordinates coordinates, PdfReadingDirection direction,
        bool explicitDirection, Action<long>? consumeWork, Action? cancellationCheck) {
        int left = FirstSpanAtOrAfter(spans, run.Start, coordinates, consumeWork, cancellationCheck) - 1;
        int right = FirstSpanAtOrAfter(spans, run.End - 0.5D, coordinates, consumeWork, cancellationCheck);
        if (left >= 0 && IsInlineScriptBoundary(run, group.Normal, spans[left], coordinates, true, direction, explicitDirection,
            consumeWork, cancellationCheck)) return spans[left];
        if (right < spans.Count && IsInlineScriptBoundary(run, group.Normal, spans[right], coordinates, false, direction, explicitDirection,
            consumeWork, cancellationCheck)) return spans[right];
        return null;
    }

    private static int FirstSpanAtOrAfter(List<PdfTextSpan> spans, double along,
        PdfTextBaselineCoordinates coordinates, Action<long>? consumeWork, Action? cancellationCheck) {
        int low = 0, high = spans.Count;
        while (low < high) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1L);
            int middle = low + (high - low) / 2;
            if (coordinates.Along(spans[middle]) < along) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    private static bool IsInlineScriptBoundary(InlineScriptRun run, double normal, PdfTextSpan anchor,
        PdfTextBaselineCoordinates coordinates, bool anchorPrecedesRun, PdfReadingDirection direction,
        bool explicitDirection, Action<long>? consumeWork, Action? cancellationCheck) {
        cancellationCheck?.Invoke();
        consumeWork?.Invoke(1L);
        TryGetInlineScriptEm(anchor, out double anchorAlongEm, out double anchorNormalEm);
        if (run.NormalEm > anchorNormalEm * 0.9D ||
            Math.Max(Math.Abs(run.MinimumNormal - normal), Math.Abs(run.MaximumNormal - normal)) > anchorNormalEm * 0.5D) return false;
        double gap = anchorPrecedesRun
            ? run.Start - coordinates.Along(anchor) - Math.Max(0D, anchor.Advance)
            : coordinates.Along(anchor) - run.End;
        if (gap < -0.5D || gap > anchorAlongEm * 0.25D) return false;
        return anchorPrecedesRun
            ? !InlineScriptBoundaryHasBreak(anchor, run.Spans[0], direction, explicitDirection)
            : !InlineScriptBoundaryHasBreak(run.Spans[run.Spans.Count - 1], anchor, direction, explicitDirection);
    }

    // Positions and advances are already in page space; FontSize remains the authored Tf value.
    // Project the transformed em across the baseline so anisotropic scaling and shear stay consistent.
    private static bool TryGetInlineScriptEm(PdfTextSpan span, out double alongEm, out double normalEm) {
        double alongScale = 1D, normalScale = 1D;
        if (span.TextToPageTransform is Matrix2D matrix) {
            alongScale = Math.Sqrt(matrix.A * matrix.A + matrix.B * matrix.B);
            normalScale = Math.Abs(matrix.A * matrix.D - matrix.B * matrix.C) / alongScale;
        }
        alongEm = span.FontSize * alongScale;
        normalEm = span.FontSize * normalScale;
        return alongEm > 0D && normalEm > 0D &&
            !double.IsNaN(alongEm) && !double.IsNaN(normalEm) &&
            !double.IsInfinity(alongEm) && !double.IsInfinity(normalEm);
    }

    // A break belongs to the following logical fragment, which may be the visually left one.
    private static bool InlineScriptBoundaryHasBreak(PdfTextSpan left, PdfTextSpan right,
        PdfReadingDirection direction, bool explicitDirection) {
        if (left.LogicalLineBreaksBefore == 0 && right.LogicalLineBreaksBefore == 0) return false;
        int order = left.ContentOrderKey is not null && right.ContentOrderKey is not null
            ? left.ContentOrderKey.CompareTo(right.ContentOrderKey) : 0;
        if (order == 0) order = left.PaintOrder.CompareTo(right.PaintOrder);
        if (order != 0) return (order < 0 ? right : left).LogicalLineBreaksBefore > 0;
        if (!explicitDirection) {
            OfficeIMO.Drawing.OfficeTextDirection leftDirection = OfficeIMO.Drawing.OfficeTextElements.ResolveBaseDirection(left.Text);
            OfficeIMO.Drawing.OfficeTextDirection rightDirection = OfficeIMO.Drawing.OfficeTextElements.ResolveBaseDirection(right.Text);
            if ((leftDirection == OfficeIMO.Drawing.OfficeTextDirection.Auto && rightDirection == OfficeIMO.Drawing.OfficeTextDirection.Auto) ||
                (leftDirection != OfficeIMO.Drawing.OfficeTextDirection.Auto && rightDirection != OfficeIMO.Drawing.OfficeTextDirection.Auto &&
                leftDirection != rightDirection)) return true;
        }
        return (direction == PdfReadingDirection.RightToLeft ? left : right).LogicalLineBreaksBefore > 0;
    }
}
