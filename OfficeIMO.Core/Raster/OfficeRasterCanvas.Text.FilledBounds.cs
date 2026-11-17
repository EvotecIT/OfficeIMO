namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // The painter and analytical inspection use the same coincident-crossing tolerance.
    private const double ContourCrossingTolerance = 1E-9D;

    // Split at vertices and edge intersections so crossing order is stable in each
    // horizontal slab. Bounds of its filled intervals occur at the slab endpoints.
    // This measures nominal flattened geometry, not raster pixel coverage.
    internal (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped)
        MeasureFilledContourBounds(IReadOnlyList<List<OfficePoint>> contours, OfficeFillRule rule,
            IReadOnlyList<OfficeTextInkClip>? clips = null) =>
        MeasureNominalFilledContourBounds(contours, rule, clips, _cancellationToken);

    // Geometry inspection retains the painter's winding, edge tolerance,
    // cancellation and bounded intersection work without allocating a surface.
    internal static (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped)
        MeasureNominalFilledContourBounds(IReadOnlyList<List<OfficePoint>> contours, OfficeFillRule rule,
            IReadOnlyList<OfficeTextInkClip>? clips = null, CancellationToken cancellationToken = default) =>
        MeasureNominalFilledContourBoundsCore(contours, rule, clips, 0, cancellationToken);

    /// <summary>Reports loss only for filled ink admitted by the prefix clips and rejected by later clips.</summary>
    internal static (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped)
        MeasureNominalAdditionalInkLoss(IReadOnlyList<List<OfficePoint>> contours, OfficeFillRule rule,
            IReadOnlyList<OfficeTextInkClip> clips, int admittedClipCount, CancellationToken cancellationToken = default) {
        ArgumentOutOfRangeException.ThrowIfNegative(admittedClipCount);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(admittedClipCount, clips.Count);
        // This overload receives filled clip geometry. Half-plane clips are applied by the
        // caller before inspection, as in the existing positioned-text inspection route.
        foreach (OfficeTextInkClip clip in clips)
            if (clip.FilledContours == null) throw new ArgumentException("Additional-loss inspection requires filled clip contours.", nameof(clips));
        return MeasureNominalFilledContourBoundsCore(contours, rule, clips, admittedClipCount, cancellationToken);
    }

    private static (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped)
        MeasureNominalFilledContourBoundsCore(IReadOnlyList<List<OfficePoint>> contours, OfficeFillRule rule,
            IReadOnlyList<OfficeTextInkClip>? clips, int admittedClipCount, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        // Diagnostics measure many small paints in a row. The working lists are borrowed per thread and returned
        // afterwards, so repeated calls reuse their capacity; a nested or failed call simply starts a new set.
        InkBoundScratch scratch = InkBoundScratch.Rent();
        try {
            return MeasureNominalFilledContourBoundsCore(contours, rule, clips, admittedClipCount, scratch, cancellationToken);
        } finally {
            InkBoundScratch.Return(scratch);
        }
    }

    private static (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped)
        MeasureNominalFilledContourBoundsCore(IReadOnlyList<List<OfficePoint>> contours, OfficeFillRule rule,
            IReadOnlyList<OfficeTextInkClip>? clips, int admittedClipCount, InkBoundScratch scratch, CancellationToken cancellationToken) {
        List<InkBoundEdge> edges = scratch.Edges;
        List<double> boundaries = scratch.Boundaries;
        List<OfficeFillRule> rules = scratch.Rules;
        rules.Add(rule);
        if (!AddContours(contours, 0)) return Unmeasured();
        if (clips != null) for (int clipIndex = 0; clipIndex < clips.Count; clipIndex++) {
            OfficeTextInkClip clip = clips[clipIndex];
            if (clip.FilledContours == null) continue;
            if (rules.Count > 64) return Unmeasured();
            rules.Add(clip.FillRule);
            if (!AddContours(clip.FilledContours, rules.Count - 1)) return Unmeasured();
        }
        long work = 4_000_000;
        for (int i = 0; i < edges.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int j = i + 1; j < edges.Count; j++) {
                if (--work < 0) return Unmeasured();
                if ((j & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                InkBoundEdge a = edges[i], b = edges[j];
                double low = Math.Max(a.LowY, b.LowY), high = Math.Min(a.HighY, b.HighY);
                if (high <= low || Math.Max(a.Start.X, a.End.X) < Math.Min(b.Start.X, b.End.X)
                    || Math.Max(b.Start.X, b.End.X) < Math.Min(a.Start.X, a.End.X)) continue;
                double atLow = a.XAt(low) - b.XAt(low), atHigh = a.XAt(high) - b.XAt(high);
                if (!IsFinite(atLow) || !IsFinite(atHigh)) return Unmeasured();
                if ((atLow < 0D && atHigh > 0D) || (atLow > 0D && atHigh < 0D)) {
                    double ratio = 1D / (1D + Math.Abs(atHigh / atLow));
                    double y = low * (1D - ratio) + high * ratio;
                    if (!IsFinite(y)) return Unmeasured();
                    if (y > low && y < high) boundaries.Add(y);
                    if (boundaries.Count > 65536) return Unmeasured();
                }
            }
        }
        boundaries.Sort();
        List<(double X, int Edge)> crossings = scratch.Crossings;
        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        bool hasInk = false, isClipped = false;
        int[] winding = scratch.Winding(rules.Count);
        for (int band = 1; band < boundaries.Count; band++) {
            cancellationToken.ThrowIfCancellationRequested();
            double low = boundaries[band - 1], high = boundaries[band];
            if (high - low <= ContourCrossingTolerance) continue;
            if ((work -= edges.Count) < 0) return Unmeasured();
            double middle = low / 2D + high / 2D;
            if (middle <= low || middle >= high) return Unmeasured();
            crossings.Clear();
            for (int edge = 0; edge < edges.Count; edge++) {
                var item = edges[edge];
                if (middle >= item.LowY && middle < item.HighY) {
                    double x = item.XAt(middle);
                    if (!IsFinite(x)) return Unmeasured();
                    crossings.Add((x, edge));
                }
            }
            crossings.Sort(CompareCrossings);
            Array.Clear(winding, 0, rules.Count);
            int index = 0, previousEdge = -1;
            double previousX = 0D;
            while (index < crossings.Count) {
                var crossing = crossings[index];
                if ((work -= rules.Count) < 0) return Unmeasured();
                bool subjectInside = IsInside(0), inside = subjectInside;
                for (int group = 1; group < rules.Count && inside; group++) inside &= IsInside(group);
                if (subjectInside && !inside && WasAdmitted() && previousEdge >= 0 && crossing.X > previousX) isClipped = true;
                if (inside && previousEdge >= 0 && crossing.X > previousX) {
                    InkBoundEdge a = edges[previousEdge], b = edges[crossing.Edge];
                    double x1 = a.XAt(low), x2 = a.XAt(high), x3 = b.XAt(low), x4 = b.XAt(high);
                    if (!IsFinite(x1) || !IsFinite(x2) || !IsFinite(x3) || !IsFinite(x4)) return Unmeasured();
                    left = Math.Min(left, Math.Min(Math.Min(x1, x2), Math.Min(x3, x4)));
                    right = Math.Max(right, Math.Max(Math.Max(x1, x2), Math.Max(x3, x4)));
                    top = Math.Min(top, low); bottom = Math.Max(bottom, high); hasInk = true;
                }
                do {
                    InkBoundEdge edge = edges[crossings[index].Edge];
                    winding[edge.Group] += rules[edge.Group] == OfficeFillRule.NonZero ? (edge.End.Y > edge.Start.Y ? 1 : -1) : 1;
                    index++;
                } while (index < crossings.Count && Math.Abs(crossings[index].X - crossing.X) <= ContourCrossingTolerance);
                previousEdge = crossing.Edge; previousX = crossing.X;
            }
        }
        return (left, top, right, bottom, hasInk, true, isClipped);

        bool WasAdmitted() {
            for (int group = 1; group <= admittedClipCount; group++) if (!IsInside(group)) return false;
            return true;
        }
        bool IsInside(int group) => rules[group] == OfficeFillRule.NonZero ? winding[group] != 0 : (winding[group] & 1) != 0;

        bool AddContours(IReadOnlyList<IReadOnlyList<OfficePoint>> source, int group) {
            for (int contourIndex = 0; contourIndex < source.Count; contourIndex++) {
                IReadOnlyList<OfficePoint> contour = source[contourIndex];
                cancellationToken.ThrowIfCancellationRequested();
                if (contour.Count < 3) continue;
                OfficePoint start = contour[contour.Count - 1];
                for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                    OfficePoint end = contour[pointIndex];
                    if (!IsFinite(start.X) || !IsFinite(start.Y) || !IsFinite(end.X) || !IsFinite(end.Y)) return false;
                    if (start.Y != end.Y) {
                        edges.Add(new InkBoundEdge(start, end, group)); boundaries.Add(start.Y); boundaries.Add(end.Y);
                        if (edges.Count > 4096) return false;
                    }
                    start = end;
                }
            }
            return true;
        }
        static (double, double, double, double, bool, bool, bool) Unmeasured() => (0D, 0D, 0D, 0D, false, false, false);
    }

    private static readonly Comparison<(double X, int Edge)> CompareCrossings = static (a, b) => a.X.CompareTo(b.X);

    private sealed class InkBoundScratch {
        // Larger buffers come from rare oversized paints; they are not kept for the thread.
        private const int RetainedEdgeCapacity = 4097;
        private const int RetainedBoundaryCapacity = 16384;

        [ThreadStatic]
        private static InkBoundScratch? _cached;

        private int[] _winding = new int[4];

        internal List<InkBoundEdge> Edges { get; } = new();
        internal List<double> Boundaries { get; } = new();
        internal List<OfficeFillRule> Rules { get; } = new();
        internal List<(double X, int Edge)> Crossings { get; } = new();

        internal int[] Winding(int count) {
            if (_winding.Length < count) _winding = new int[Math.Max(count, _winding.Length * 2)];
            return _winding;
        }

        internal static InkBoundScratch Rent() {
            InkBoundScratch? scratch = _cached;
            _cached = null;
            return scratch ?? new InkBoundScratch();
        }

        internal static void Return(InkBoundScratch scratch) {
            if (scratch.Edges.Capacity > RetainedEdgeCapacity || scratch.Boundaries.Capacity > RetainedBoundaryCapacity
                || scratch.Crossings.Capacity > RetainedEdgeCapacity) return;
            scratch.Edges.Clear();
            scratch.Boundaries.Clear();
            scratch.Rules.Clear();
            scratch.Crossings.Clear();
            _cached = scratch;
        }
    }

    private readonly struct InkBoundEdge {
        internal InkBoundEdge(OfficePoint start, OfficePoint end, int group) { Start = start; End = end; Group = group; }
        internal int Group { get; }
        internal OfficePoint Start { get; }
        internal OfficePoint End { get; }
        internal double LowY => Math.Min(Start.Y, End.Y);
        internal double HighY => Math.Max(Start.Y, End.Y);
        internal double XAt(double y) {
            double ratio = (y - Start.Y) / (End.Y - Start.Y);
            return Start.X * (1D - ratio) + End.X * ratio;
        }
    }
}
