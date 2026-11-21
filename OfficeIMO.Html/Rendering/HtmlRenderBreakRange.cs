namespace OfficeIMO.Html;

/// <summary>
/// A closed source-coordinate interval in which a fragment may end without
/// cutting content. Paragraph and authored keep constraints still apply.
/// </summary>
internal sealed record HtmlRenderBreakRange(double Start, double End) {
    internal HtmlRenderBreakRange Translate(double offset) => new(Start + offset, End + offset);

    internal static IReadOnlyList<HtmlRenderBreakRange> Normalize(
        IEnumerable<HtmlRenderBreakRange> source, double extent) {
        if (source is IReadOnlyCollection<HtmlRenderBreakRange> { Count: 0 }) return Array.Empty<HtmlRenderBreakRange>();

        var candidates = new List<HtmlRenderBreakRange>();
        bool ordered = true;
        foreach (HtmlRenderBreakRange range in source) {
            if (!IsFinite(range.Start) || !IsFinite(range.End)
                || range.End < range.Start || range.End < 0D || range.Start > extent) continue;
            var clamped = new HtmlRenderBreakRange(Math.Max(0D, range.Start), Math.Min(extent, range.End));
            if (ordered && candidates.Count > 0) {
                HtmlRenderBreakRange last = candidates[candidates.Count - 1];
                int byStart = last.Start.CompareTo(clamped.Start);
                ordered = byStart < 0 || byStart == 0 && last.End.CompareTo(clamped.End) <= 0;
            }
            candidates.Add(clamped);
        }
        // Ranges almost always arrive in order; the stable sort is kept for the others.
        IEnumerable<HtmlRenderBreakRange> sorted = ordered
            ? candidates
            : candidates.OrderBy(range => range.Start).ThenBy(range => range.End);
        var result = new List<HtmlRenderBreakRange>();
        foreach (HtmlRenderBreakRange range in sorted) {
            if (result.Count > 0 && range.Start <= result[result.Count - 1].End) {
                HtmlRenderBreakRange previous = result[result.Count - 1];
                result[result.Count - 1] = new(previous.Start, Math.Max(previous.End, range.End));
            } else result.Add(range);
        }
        return result.AsReadOnly();
    }

    internal static IReadOnlyList<HtmlRenderBreakRange> Intersect(
        IEnumerable<HtmlRenderBreakRange> left, IEnumerable<HtmlRenderBreakRange> right) {
        IReadOnlyList<HtmlRenderBreakRange> first = Normalize(left, double.MaxValue);
        IReadOnlyList<HtmlRenderBreakRange> second = Normalize(right, double.MaxValue);
        var result = new List<HtmlRenderBreakRange>();
        int a = 0, b = 0;
        while (a < first.Count && b < second.Count) {
            double start = Math.Max(first[a].Start, second[b].Start);
            double end = Math.Min(first[a].End, second[b].End);
            if (end >= start) result.Add(new(start, end));
            if (first[a].End < second[b].End) a++; else b++;
        }
        return result.AsReadOnly();
    }

    /// <summary>Subtracts open carrier interiors; adjacent carriers retain their shared legal endpoint.</summary>
    internal static IReadOnlyList<HtmlRenderBreakRange> Exclude(
        IEnumerable<HtmlRenderBreakRange> source, IEnumerable<HtmlRenderBreakRange> forbidden) {
        IReadOnlyList<HtmlRenderBreakRange> ranges = Normalize(source, double.MaxValue);
        // Do not merge touching forbidden ranges: the shared endpoint is legal.
        var exclusions = new List<HtmlRenderBreakRange>();
        foreach (HtmlRenderBreakRange range in forbidden.Where(range => IsFinite(range.Start)
            && IsFinite(range.End) && range.End > range.Start).OrderBy(range => range.Start)) {
            if (exclusions.Count > 0 && range.Start < exclusions[exclusions.Count - 1].End) {
                HtmlRenderBreakRange previous = exclusions[exclusions.Count - 1];
                exclusions[exclusions.Count - 1] = new(previous.Start, Math.Max(previous.End, range.End));
            } else exclusions.Add(range);
        }
        var result = new List<HtmlRenderBreakRange>();
        int first = 0;
        foreach (HtmlRenderBreakRange range in ranges) {
            double start = range.Start;
            while (first < exclusions.Count && exclusions[first].End < start) first++;
            for (int i = first; i < exclusions.Count && exclusions[i].Start <= range.End; i++) {
                HtmlRenderBreakRange excluded = exclusions[i];
                if (excluded.End <= start) continue;
                if (excluded.Start >= start) result.Add(new(start, Math.Min(range.End, excluded.Start)));
                start = Math.Max(start, excluded.End);
                if (start > range.End) break;
            }
            if (start <= range.End) result.Add(new(start, range.End));
        }
        return result.AsReadOnly();
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
