using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.IWork.Internal;

/// <summary>Builds disjoint sparse edge intervals, choosing the latest native edit on each shared segment.</summary>
internal sealed class IWorkTableBorderOverrides {
    private readonly Dictionary<(bool Horizontal, int Line), List<Run>> _runs = new();
    private readonly Dictionary<(bool Horizontal, int Line), List<Segment>> _segments = new();
    private bool _hasDeclarations;
    internal bool FullyReconstructed { get; private set; } = true;
    internal bool LocationsAreComplete { get; private set; } = true;
    internal bool HasDeclarations => _hasDeclarations || !LocationsAreComplete;

    internal void RejectLocations() { LocationsAreComplete = FullyReconstructed = false; }

    internal void Add(bool horizontal, int line, int origin, int length, uint order, IWorkBorderDeclaration stroke) {
        var key = (horizontal, line);
        if (!_runs.TryGetValue(key, out var runs)) _runs.Add(key, runs = new List<Run>());
        runs.Add(new Run(origin, checked(origin + length), order, stroke));
        _hasDeclarations = true;
        if (!stroke.IsSupported) FullyReconstructed = false;
    }

    internal IWorkBorderDeclaration? Get(bool horizontal, int line, int position) {
        if (!LocationsAreComplete) return IWorkBorderDeclaration.Invalid;
        if (!_segments.TryGetValue((horizontal, line), out var segments)) return null;
        int low = 0, high = segments.Count - 1;
        while (low <= high) {
            int middle = low + (high - low) / 2;
            Segment segment = segments[middle];
            if (position < segment.Start) high = middle - 1;
            else if (position >= segment.End) low = middle + 1;
            else return segment.Stroke;
        }
        return null;
    }

    internal bool CoversAllEdges(int rows, int columns, CancellationToken cancellationToken) {
        if (!FullyReconstructed || !LocationsAreComplete || rows <= 0 || columns <= 0
            || _segments.Count != (long)rows + columns + 2) return false;
        // Check only the bounded sparse representation, never allocate or walk the cell grid.
        foreach (var pair in _segments) {
            cancellationToken.ThrowIfCancellationRequested();
            int lastLine = pair.Key.Horizontal ? rows : columns;
            if (pair.Key.Line < 0 || pair.Key.Line > lastLine) return false;
            int extent = pair.Key.Horizontal ? columns : rows;
            int end = 0;
            foreach (var segment in pair.Value) {
                cancellationToken.ThrowIfCancellationRequested();
                if (segment.Start != end || !segment.Stroke.IsSupported) return false;
                end = segment.End;
            }
            if (end != extent) return false;
        }
        return true;
    }

    internal void Seal(CancellationToken cancellationToken) {
        foreach (var pair in _runs) {
            cancellationToken.ThrowIfCancellationRequested();
            var events = new SortedDictionary<int, List<(Run Run, bool Start)>>();
            foreach (var run in pair.Value) {
                AddEvent(run.Start, run, true); AddEvent(run.End, run, false);
            }
            var active = new Dictionary<uint, Dictionary<(bool Supported, OfficeBorderSide? Side), int>>();
            var orders = new SortedSet<uint>();
            var segments = new List<Segment>();
            int? previous = null;
            foreach (var point in events) {
                cancellationToken.ThrowIfCancellationRequested();
                if (previous.HasValue && previous.Value < point.Key && orders.Count > 0) {
                    var styles = active[orders.Max];
                    var style = styles.First().Key;
                    IWorkBorderDeclaration stroke = styles.Count == 1
                        ? new IWorkBorderDeclaration(style.Side, style.Supported) : IWorkBorderDeclaration.Invalid;
                    if (!stroke.IsSupported) FullyReconstructed = false;
                    if (segments.Count > 0 && segments[segments.Count - 1].End == previous.Value
                        && segments[segments.Count - 1].Stroke.IsSupported == stroke.IsSupported
                        && Nullable.Equals(segments[segments.Count - 1].Stroke.Side, stroke.Side))
                        segments[segments.Count - 1].End = point.Key;
                    else segments.Add(new Segment(previous.Value, point.Key, stroke));
                }
                foreach (var change in point.Value) {
                    var run = change.Run;
                    var key = (run.Stroke.IsSupported, run.Stroke.Side);
                    if (change.Start) {
                        if (!active.TryGetValue(run.Order, out var styles)) {
                            active.Add(run.Order, styles = new()); orders.Add(run.Order);
                        }
                        styles.TryGetValue(key, out int count); styles[key] = count + 1;
                    } else {
                        var styles = active[run.Order];
                        if (styles[key] == 1) styles.Remove(key); else styles[key]--;
                        if (styles.Count == 0) { active.Remove(run.Order); orders.Remove(run.Order); }
                    }
                }
                previous = point.Key;
            }
            _segments.Add(pair.Key, segments);

            void AddEvent(int position, Run run, bool start) {
                if (!events.TryGetValue(position, out var values)) events.Add(position, values = new());
                values.Add((run, start));
            }
        }
        _runs.Clear();
    }

    private sealed class Run(int start, int end, uint order, IWorkBorderDeclaration stroke) {
        internal int Start { get; } = start;
        internal int End { get; } = end;
        internal uint Order { get; } = order;
        internal IWorkBorderDeclaration Stroke { get; } = stroke;
    }
    private sealed class Segment(int start, int end, IWorkBorderDeclaration stroke) {
        internal int Start { get; } = start;
        internal int End { get; set; } = end;
        internal IWorkBorderDeclaration Stroke { get; } = stroke;
    }
}
