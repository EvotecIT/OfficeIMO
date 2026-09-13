namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Proposes deterministic delays and optional remaining-work interruptions to resolve resource overloads. Preserves recorded actuals.</summary>
    public ProjectLevelingResult CalculateLeveling(ProjectLevelingOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed(); options ??= new ProjectLevelingOptions();
        if (options.MaxIterations < 1 || options.MaxIterations > 1_000_000 || options.MaxDelayDays < 0 || options.MaxDelayDays > 366000)
            throw new ArgumentOutOfRangeException(nameof(options));
        if (options.MaxSplits < 1 || options.MaxSplits > 100000) throw new ArgumentOutOfRangeException(nameof(options));
        if (options.ScheduleOptions == null || !options.ScheduleOptions.CalculateAssignments)
            throw new ArgumentException("Leveling requires independent assignment calculation settings.", nameof(options));
        return ProjectLeveler.Calculate(this, options, cancellationToken);
    }
    /// <summary>Applies a complete leveling proposal to its unchanged originating document.</summary>
    public void ApplyLeveling(ProjectLevelingResult result, CancellationToken cancellationToken = default) {
        EnsureMutable();
        if (result == null) throw new ArgumentNullException(nameof(result));
        result.Report.ThrowIfErrors();
        if (result.Schedule.Document != this || result.Schedule.ModelRevision != Revision || _batchDepth != 0)
            throw new InvalidOperationException("The leveling proposal belongs to another document or a stale/in-progress revision.");
        var updates = new List<(ProjectTask Task, ProjectDuration Delay)>();
        foreach (var move in result.Moves) {
            cancellationToken.ThrowIfCancellationRequested(); var task = Tasks.GetByUid(move.TaskUid);
            var scheduled = result.Schedule.Tasks.Single(t => t.TaskUid == move.TaskUid);
            decimal delay = (scheduled.ForwardAnchor.Ticks - scheduled.BeforeLevelingAnchor.Ticks) / (decimal)TimeSpan.TicksPerMinute;
            if (delay < 0 || delay * 10m % 1m != 0) throw new InvalidOperationException("The proposed leveling delay cannot be represented in whole tenths of a minute.");
            updates.Add((task, new ProjectDuration(delay, ProjectDurationUnit.Minute, elapsed: true)));
        }
        ApplyScheduleCore(result.Schedule, cancellationToken, () => {
            foreach (var update in updates) update.Task.LevelingDelay = update.Delay;
            var splitTasks = new HashSet<int>(result.Splits.Select(s => s.TaskUid));
            foreach (var assignment in Assignments.Where(a => a.Task != null && splitTasks.Contains(a.Task.Uid) && a.TimephasedData.Any(v => v.Type == 1 || v.Type == 2)))
                assignment.WorkContour = ProjectWorkContour.Custom;
        });
    }
}
