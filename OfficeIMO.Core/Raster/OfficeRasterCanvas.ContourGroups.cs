namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Fills the shape as separate groups of contours that share no device pixel, such as the glyphs of a
    /// text run. Every pixel still receives the coverage of all contours that touch it, so winding is
    /// unchanged; each group splits its sub-scanlines only at its own vertex heights.
    /// </summary>
    internal void FillDisjointContourGroups(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule,
        Func<double, double, OfficeColor> paint, IReadOnlyList<IReadOnlyList<OfficePoint>>? unionContours = null) {
        if (contours == null || contours.Count == 0) return;
        int total = contours.Count + (unionContours?.Count ?? 0);
        if (total < 2 || !TryGroupDisjointContours(contours, unionContours, out ContourGroupEntry[] entries, out int[] groupStarts)) {
            FillContourPaint(contours, fillRule, paint, unionContours);
            return;
        }
        double scaleX = 1D, scaleY = 1D;
        bool scaled = HasCoordinateScale;
        if (scaled) {
            scaleX = CoordinateScaleX;
            scaleY = CoordinateScaleY;
            if (paint.Target is not SolidContourPaint) {
                Func<double, double, OfficeColor> localPaint = paint;
                paint = (x, y) => localPaint(x / scaleX, y / scaleY);
            }
        }
        var primary = new ContourGroupView(contours);
        ContourGroupView? union = unionContours != null ? new ContourGroupView(unionContours) : null;
        ContourCoverageWorkspace workspace = TakeContourCoverageWorkspace();
        try {
            for (int group = 0; group + 1 < groupStarts.Length; group++) {
                _cancellationToken.ThrowIfCancellationRequested();
                int start = groupStarts[group], end = groupStarts[group + 1];
                // Source order inside a group keeps the edge order the single fill would collect.
                Array.Sort(entries, start, end - start, ContourGroupEntryOrder.Instance);
                primary.Reset();
                union?.Reset();
                for (int index = start; index < end; index++) {
                    ContourGroupEntry entry = entries[index];
                    if (entry.Union) union!.Add(entry.Index);
                    else primary.Add(entry.Index);
                }
                if (primary.Count == 0 && (union == null || union.Count == 0)) continue;
                FillContourPaint(primary, fillRule, paint, union != null && union.Count > 0 ? union : null, workspace, null, scaled, scaleX, scaleY);
            }
        } finally {
            ReturnContourCoverageWorkspace(workspace);
        }
    }

    // Groups contours along the run's longer axis: a contour joins the current group when its first
    // device pixel on that axis is at or before the group's last one. Groups that share no pixel
    // column (or row) cannot share a pixel. Returns false when the shape forms a single group.
    private bool TryGroupDisjointContours(IReadOnlyList<IReadOnlyList<OfficePoint>> contours,
        IReadOnlyList<IReadOnlyList<OfficePoint>>? unionContours, out ContourGroupEntry[] entries, out int[] groupStarts) {
        groupStarts = Array.Empty<int>();
        double scaleX = CoordinateScaleX, scaleY = CoordinateScaleY;
        int total = contours.Count + (unionContours?.Count ?? 0);
        entries = new ContourGroupEntry[total];
        int count = 0;
        double minX = double.PositiveInfinity, maxX = double.NegativeInfinity;
        double minY = double.PositiveInfinity, maxY = double.NegativeInfinity;
        for (int list = 0; list < 2; list++) {
            IReadOnlyList<IReadOnlyList<OfficePoint>>? source = list == 0 ? contours : unionContours;
            if (source == null) continue;
            for (int index = 0; index < source.Count; index++) {
                IReadOnlyList<OfficePoint> contour = source[index];
                // The filler skips contours with fewer than three points, so they need no group.
                if (contour == null || contour.Count < 3) continue;
                double left = double.PositiveInfinity, right = double.NegativeInfinity;
                double top = double.PositiveInfinity, bottom = double.NegativeInfinity;
                for (int point = 0; point < contour.Count; point++) {
                    OfficePoint value = contour[point];
                    double x = value.X * scaleX, y = value.Y * scaleY;
                    if (!IsFinite(x) || !IsFinite(y)) return false;
                    if (x < left) left = x;
                    if (x > right) right = x;
                    if (y < top) top = y;
                    if (y > bottom) bottom = y;
                }
                entries[count++] = new ContourGroupEntry(left, right, top, bottom, index, list == 1);
                minX = Math.Min(minX, left); maxX = Math.Max(maxX, right);
                minY = Math.Min(minY, top); maxY = Math.Max(maxY, bottom);
            }
        }
        if (count < 2) return false;
        bool horizontal = maxX - minX >= maxY - minY;
        Array.Sort(entries, 0, count, horizontal ? ContourGroupEntryAxisOrder.Horizontal : ContourGroupEntryAxisOrder.Vertical);
        var starts = new List<int> { 0 };
        // Pixels touched on the axis run from floor(low) to ceil(high) - 1, as in the filler.
        double groupEnd = Math.Ceiling(entries[0].High(horizontal));
        for (int index = 1; index < count; index++) {
            ContourGroupEntry entry = entries[index];
            if (Math.Floor(entry.Low(horizontal)) >= groupEnd) starts.Add(index);
            groupEnd = Math.Max(groupEnd, Math.Ceiling(entry.High(horizontal)));
        }
        if (starts.Count == 1) return false;
        starts.Add(count);
        groupStarts = starts.ToArray();
        return true;
    }

    private readonly struct ContourGroupEntry {
        internal ContourGroupEntry(double left, double right, double top, double bottom, int index, bool union) {
            Left = left; Right = right; Top = top; Bottom = bottom; Index = index; Union = union;
        }
        internal double Left { get; }
        internal double Right { get; }
        internal double Top { get; }
        internal double Bottom { get; }
        internal int Index { get; }
        internal bool Union { get; }
        internal double Low(bool horizontal) => horizontal ? Left : Top;
        internal double High(bool horizontal) => horizontal ? Right : Bottom;
    }

    private sealed class ContourGroupEntryAxisOrder : IComparer<ContourGroupEntry> {
        internal static readonly ContourGroupEntryAxisOrder Horizontal = new ContourGroupEntryAxisOrder(true);
        internal static readonly ContourGroupEntryAxisOrder Vertical = new ContourGroupEntryAxisOrder(false);
        private readonly bool _horizontal;
        private ContourGroupEntryAxisOrder(bool horizontal) { _horizontal = horizontal; }
        public int Compare(ContourGroupEntry x, ContourGroupEntry y) {
            int order = x.Low(_horizontal).CompareTo(y.Low(_horizontal));
            return order != 0 ? order : ContourGroupEntryOrder.Instance.Compare(x, y);
        }
    }

    private sealed class ContourGroupEntryOrder : IComparer<ContourGroupEntry> {
        internal static readonly ContourGroupEntryOrder Instance = new ContourGroupEntryOrder();
        public int Compare(ContourGroupEntry x, ContourGroupEntry y) {
            int order = x.Union.CompareTo(y.Union);
            return order != 0 ? order : x.Index.CompareTo(y.Index);
        }
    }

    /// <summary>A reusable view of selected contours of one source list.</summary>
    private sealed class ContourGroupView : IReadOnlyList<IReadOnlyList<OfficePoint>> {
        private readonly IReadOnlyList<IReadOnlyList<OfficePoint>> _source;
        private int[] _indices = new int[8];

        internal ContourGroupView(IReadOnlyList<IReadOnlyList<OfficePoint>> source) { _source = source; }

        public int Count { get; private set; }
        public IReadOnlyList<OfficePoint> this[int index] => _source[_indices[index]];

        internal void Reset() => Count = 0;

        internal void Add(int index) {
            if (Count == _indices.Length) Array.Resize(ref _indices, _indices.Length * 2);
            _indices[Count++] = index;
        }

        public IEnumerator<IReadOnlyList<OfficePoint>> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
