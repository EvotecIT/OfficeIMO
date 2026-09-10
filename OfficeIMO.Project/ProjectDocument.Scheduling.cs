namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Calculates a schedule without changing stored dates, actuals, costs, or source bytes. Unsupported combinations return diagnostics.</summary>
    public ProjectScheduleResult CalculateSchedule(ProjectScheduleOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("Finish the update scope before calculating a revision-bound schedule.");
        options ??= new ProjectScheduleOptions();
        if (options.MaxCalendarDays < 1 || options.MaxCalendarDays > 366000 || options.CriticalSlackMinutes < 0)
            throw new ArgumentOutOfRangeException(nameof(options));
        return new ProjectScheduler(this, options, cancellationToken).Calculate();
    }
    /// <summary>Applies an error-free schedule produced for this exact document revision. No work/cost/actual recalculation is implied.</summary>
    public void ApplySchedule(ProjectScheduleResult result, CancellationToken cancellationToken = default) {
        EnsureMutable();
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (_batchDepth != 0 || result.Document != this || result.ModelRevision != Revision)
            throw new InvalidOperationException("The schedule belongs to another document or a stale/in-progress revision.");
        result.Report.ThrowIfErrors();
        cancellationToken.ThrowIfCancellationRequested();
        // All potentially failing validation precedes mutation. Cancellation is intentionally observed only before this atomic model update.
        foreach (var item in result.Tasks) if (!TaskIndex.ContainsKey(item.TaskUid)) throw new InvalidOperationException("A scheduled task is no longer attached.");
        var dates = result.Tasks.ToDictionary(t => t.TaskUid);
        using (BeginUpdate()) {
            if (!Settings.ScheduleFromStart.HasValue) Settings.ScheduleFromStart = true;
            foreach (var item in result.Tasks) {
                var task = TaskIndex[item.TaskUid]; task.Start = item.Start; task.Finish = item.Finish; task.IsCritical = item.IsCritical;
                task.Duration = item.Duration;
                if (!task.IsSummary && task.IsManual != true) task.RemainingDuration = item.Duration;
                task.EarlyStart = item.EarlyStart; task.EarlyFinish = item.EarlyFinish; task.LateStart = item.LateStart; task.LateFinish = item.LateFinish;
                task.TotalSlackMinutes = item.TotalSlackMinutes; task.FreeSlackMinutes = item.FreeSlackMinutes;
            }
            foreach (var assignment in Assignments) {
                if (assignment.Task == null || !dates.TryGetValue(assignment.Task.Uid, out var item)) continue;
                // Existing conflicting assignment dates are rejected by calculation. Newly authored
                // assignments need explicit endpoints so an importer does not choose the project origin.
                if (!assignment.Start.HasValue) assignment.Start = item.Start;
                if (!assignment.Finish.HasValue) assignment.Finish = item.Finish;
            }
        }
        IsScheduleStale = false;
    }
    /// <summary>Calculates and explicitly applies task dates. Throws before mutation when calculation reports errors.</summary>
    public ProjectScheduleResult Recalculate(ProjectScheduleOptions? options = null, CancellationToken cancellationToken = default) {
        var result = CalculateSchedule(options, cancellationToken); ApplySchedule(result, cancellationToken); return result;
    }
}
