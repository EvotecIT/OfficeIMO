using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Regular work demand and dated availability for one uninterrupted resource interval.</summary>
public sealed class ProjectResourceAllocationInterval {
    internal ProjectResourceAllocationInterval(int uid, DateTime start, DateTime finish, decimal units, decimal capacity, IEnumerable<int> assignments) {
        ResourceUid = uid; Start = start; Finish = finish; Units = units; Capacity = capacity;
        AssignmentUids = new ReadOnlyCollection<int>(assignments.OrderBy(a => a).ToArray());
    }
    /// <summary>Resource identity.</summary>
    public int ResourceUid { get; }
    /// <summary>Inclusive interval start.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive interval finish.</summary>
    public DateTime Finish { get; }
    /// <summary>Sum of concurrent regular work divided by working minutes.</summary>
    public decimal Units { get; }
    /// <summary>Available units from the resource's dated availability table, or MaxUnits.</summary>
    public decimal Capacity { get; }
    /// <summary>Demand exceeding available units; zero when the interval fits.</summary>
    public decimal ExcessUnits => Math.Max(0m, Units - Capacity);
    /// <summary>Assignments contributing positive regular work in this interval.</summary>
    public IReadOnlyList<int> AssignmentUids { get; }
}

/// <summary>Immutable capacity analysis of an explicitly calculated assignment schedule.</summary>
public sealed class ProjectResourceAllocationResult {
    internal ProjectResourceAllocationResult(long revision, IEnumerable<ProjectResourceAllocationInterval> intervals) {
        ModelRevision = revision; Intervals = new ReadOnlyCollection<ProjectResourceAllocationInterval>(intervals.ToArray());
        Overallocations = new ReadOnlyCollection<ProjectResourceAllocationInterval>(Intervals.Where(i => i.ExcessUnits > .000000001m).ToArray());
    }
    /// <summary>Document revision analyzed.</summary>
    public long ModelRevision { get; }
    /// <summary>Occupied regular-work intervals split at changes in demand or availability.</summary>
    public IReadOnlyList<ProjectResourceAllocationInterval> Intervals { get; }
    /// <summary>Intervals whose demand exceeds availability, with a tolerance of one billionth of a unit.</summary>
    public IReadOnlyList<ProjectResourceAllocationInterval> Overallocations { get; }
}

internal static class ProjectResourceAllocation {
    private readonly struct Change {
        internal Change(int uid, decimal delta) { Uid = uid; Delta = delta; }
        internal int Uid { get; }
        internal decimal Delta { get; }
    }
    internal static ProjectResourceAllocationResult Analyze(ProjectDocument document, ProjectScheduleResult schedule, int limit, CancellationToken token) {
        var results = new List<ProjectResourceAllocationInterval>();
        foreach (var group in schedule.Assignments.GroupBy(a => a.ResourceUid).OrderBy(g => g.Key)) {
            token.ThrowIfCancellationRequested();
            var resource = document.Resources.GetByUid(group.Key);
            if (resource.Type != ProjectResourceType.Work) continue;
            AnalyzeResource(resource, group.Select(a => (a.AssignmentUid, a.Intervals)), results, limit, token);
        }
        return new ProjectResourceAllocationResult(schedule.ModelRevision, results);
    }

    internal static void AnalyzeResource(ProjectResource resource, IEnumerable<(int AssignmentUid, IReadOnlyList<ProjectAssignmentInterval> Intervals)> assignments,
        List<ProjectResourceAllocationInterval> results, int limit, CancellationToken token) {
            var events = new SortedDictionary<DateTime, List<Change>>();
            void Add(DateTime at, int uid, decimal amount) {
                if (!events.TryGetValue(at, out var changes)) {
                    events.Add(at, changes = new List<Change>());
                    if (events.Count > limit * 2L) throw new InvalidOperationException("Resource allocation analysis exceeds MaxIntervals.");
                }
                changes.Add(new Change(uid, amount));
            }
            foreach (var assignment in assignments) foreach (var interval in assignment.Intervals) {
                token.ThrowIfCancellationRequested();
                if (interval.Units <= 0) continue;
                Add(interval.Start, assignment.AssignmentUid, interval.Units); Add(interval.Finish, assignment.AssignmentUid, -interval.Units);
            }
            if (events.Count == 0) return;
            var timeline = new ProjectResourceTimeline(resource);
            foreach (var at in timeline.CapacityBoundaries(events.First().Key, events.Last().Key)) Add(at, -1, 0m);
            var active = new Dictionary<int, decimal>(); DateTime? previous = null; decimal total = 0m;
            foreach (var pair in events) {
                token.ThrowIfCancellationRequested();
                if (previous.HasValue && pair.Key > previous && total > .000000001m) {
                    results.Add(new ProjectResourceAllocationInterval(resource.Uid, previous.Value, pair.Key, total, timeline.Capacity(previous.Value), active.Keys));
                    if (results.Count > limit) throw new InvalidOperationException("Resource allocation analysis exceeds MaxIntervals.");
                }
                foreach (var change in pair.Value) {
                    if (change.Delta == 0) continue;
                    active.TryGetValue(change.Uid, out var current); current += change.Delta; total += change.Delta;
                    if (Math.Abs(current) <= .000000001m) active.Remove(change.Uid); else active[change.Uid] = current;
                }
                previous = pair.Key;
            }
    }
}
