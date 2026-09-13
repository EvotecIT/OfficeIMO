namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>
    /// Adds a finite expanded recurring summary from strictly increasing local occurrence dates.
    /// Children use start-no-earlier-than constraints and the supplied duration; normal calendar scheduling determines their dates.
    /// XML retains the occurrences and recurring markers. Microsoft Project imports these as ordinary tasks and does not reconstruct an editable recurrence rule.
    /// Enumeration, cancellation, ownership and bounds are checked before adding any task.
    /// </summary>
    public ProjectTask AddRecurringTask(string name, IEnumerable<DateTime> occurrenceStarts, ProjectDuration duration,
        ProjectTask? parent = null, ProjectCalendar? calendar = null, int maxOccurrences = 10000,
        CancellationToken cancellationToken = default) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (occurrenceStarts == null) throw new ArgumentNullException(nameof(occurrenceStarts));
        if (duration.Value < 0) throw new ArgumentOutOfRangeException(nameof(duration));
        if (maxOccurrences < 1 || maxOccurrences > 100000) throw new ArgumentOutOfRangeException(nameof(maxOccurrences));
        EnsureMutable(); CheckMember(parent); CheckMember(calendar);
        if (parent?.Uid == 0) throw new ArgumentException("The reserved project summary cannot own outline tasks.", nameof(parent));
        var starts = new List<DateTime>();
        foreach (var start in occurrenceStarts) {
            cancellationToken.ThrowIfCancellationRequested();
            if (starts.Count == maxOccurrences) throw new InvalidOperationException("The recurring task exceeds the occurrence limit.");
            if (start.Kind != DateTimeKind.Unspecified)
                throw new ArgumentException("Occurrence dates must be local project dates with unspecified DateTime kind.", nameof(occurrenceStarts));
            if (starts.Count != 0 && start <= starts[starts.Count - 1])
                throw new ArgumentException("Occurrence dates must be strictly increasing and distinct.", nameof(occurrenceStarts));
            starts.Add(start);
        }
        if (starts.Count == 0) throw new ArgumentException("At least one occurrence is required.", nameof(occurrenceStarts));
        cancellationToken.ThrowIfCancellationRequested();
        // A caller-supplied iterator can change the document during enumeration.
        EnsureMutable(); CheckMember(parent); CheckMember(calendar);
        if ((long)Math.Max(_taskUid, TaskIndex.Count == 0 ? 0 : TaskIndex.Keys.Max()) + starts.Count + 2 > int.MaxValue)
            throw new InvalidOperationException("There are not enough task UIDs for the recurring series.");
        using (BeginUpdate()) {
            var summary = (parent?.Children ?? Tasks).AddSummary(name);
            summary.IsRecurring = true; summary.IsManual = false;
            foreach (var start in starts) {
                var child = summary.Children.Add(name);
                child.IsRecurring = true; child.IsManual = false; child.Duration = duration;
                child.Calendar = calendar; child.ConstraintType = ProjectConstraintType.StartNoEarlierThan;
                child.ConstraintDate = start;
            }
            return summary;
        }
    }
}
