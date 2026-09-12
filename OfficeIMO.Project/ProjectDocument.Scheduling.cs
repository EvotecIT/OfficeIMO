namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Calculates a schedule without changing stored dates, actuals, costs, or source bytes. Unsupported combinations return diagnostics.</summary>
    public ProjectScheduleResult CalculateSchedule(ProjectScheduleOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("Finish the update scope before calculating a revision-bound schedule.");
        options ??= new ProjectScheduleOptions();
        if (options.MaxCalendarDays < 1 || options.MaxCalendarDays > 366000 || options.CriticalSlackMinutes < 0 || options.MaxIntervals < 1 ||
            options.MaxExternalProjects < 1 || options.MaxExternalProjects > 1024 || options.MaxExternalDepth < 1 || options.MaxExternalDepth > 64 || options.MaxExternalTasks < 1)
            throw new ArgumentOutOfRangeException(nameof(options));
        if (!options.CalculateAssignments && (options.RedistributeEffortDrivenWork || options.RescheduleRemainingAfterStatusDate || options.RecalculateActualCosts))
            throw new ArgumentException("Assignment work/progress/cost options require CalculateAssignments.", nameof(options));
        if (options.RescheduleRemainingAfterStatusDate && !Settings.StatusDate.HasValue)
            throw new ArgumentException("Rescheduling remaining work requires an explicit StatusDate.", nameof(options));
        return new ProjectScheduler(this, options, cancellationToken).Calculate();
    }
    /// <summary>Applies an error-free schedule produced for this exact document revision, including assignment calculations when explicitly requested.</summary>
    public void ApplySchedule(ProjectScheduleResult result, CancellationToken cancellationToken = default) => ApplyScheduleCore(result, cancellationToken);
    private void ApplyScheduleCore(ProjectScheduleResult result, CancellationToken cancellationToken, Action? additionalUpdates = null) {
        EnsureMutable();
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (_batchDepth != 0 || result.Document != this || result.ModelRevision != Revision)
            throw new InvalidOperationException("The schedule belongs to another document or a stale/in-progress revision.");
        result.Report.ThrowIfErrors();
        foreach (var source in result.ExternalSources) source.ValidateCurrent();
        cancellationToken.ThrowIfCancellationRequested();
        // All potentially failing validation precedes mutation. Cancellation is intentionally observed only before this atomic model update.
        foreach (var item in result.Tasks) if (!TaskIndex.ContainsKey(item.TaskUid)) throw new InvalidOperationException("A scheduled task is no longer attached.");
        var dates = result.Tasks.ToDictionary(t => t.TaskUid);
        var assignmentUpdates = PrepareAssignmentUpdates(result);
        var resourceUpdates = PrepareResourceUpdates(result);
        var progressDurations = result.Tasks.Where(t => t.Calculation != null).ToDictionary(t => t.TaskUid, t => {
            decimal scale = ProjectXmlValue.MinutesPerUnit(t.Duration.Unit, t.Duration.IsElapsed, this);
            return (Actual: new ProjectDuration(t.Calculation!.ActualDuration.Value / scale, t.Duration.Unit, t.Duration.IsElapsed, t.Duration.IsEstimated),
                Remaining: new ProjectDuration(t.Calculation.RemainingDuration.Value / scale, t.Duration.Unit, t.Duration.IsElapsed, t.Duration.IsEstimated));
        });
        using (BeginUpdate()) {
            if (result.CalculatedAssignments) Settings.ExternallyEdited = false;
            if (!Settings.ScheduleFromStart.HasValue) Settings.ScheduleFromStart = true;
            foreach (var item in result.Tasks) {
                var task = TaskIndex[item.TaskUid]; task.Start = item.Start; task.Finish = item.Finish; task.IsCritical = item.IsCritical;
                task.Duration = item.Duration;
                if (item.Calculation is ProjectTaskWorkSchedule calculation) {
                    task.Work = calculation.Work; task.ActualWork = calculation.ActualWork; task.RemainingWork = calculation.RemainingWork;
                    task.ActualDuration = progressDurations[item.TaskUid].Actual; task.RemainingDuration = progressDurations[item.TaskUid].Remaining;
                    task.PercentComplete = calculation.PercentComplete; task.PercentWorkComplete = calculation.PercentWorkComplete;
                    task.Cost = calculation.Cost; task.ActualCost = calculation.ActualCost; task.RemainingCost = calculation.RemainingCost;
                } else if (!task.IsSummary && task.IsManual != true) task.RemainingDuration = item.Duration;
                task.EarlyStart = item.EarlyStart; task.EarlyFinish = item.EarlyFinish; task.LateStart = item.LateStart; task.LateFinish = item.LateFinish;
                task.TotalSlackMinutes = item.TotalSlackMinutes; task.FreeSlackMinutes = item.FreeSlackMinutes;
            }
            foreach (var update in assignmentUpdates) update.Apply();
            foreach (var update in resourceUpdates) update.Apply();
            additionalUpdates?.Invoke();
            foreach (var assignment in Assignments) {
                if (assignment.Task == null || !dates.TryGetValue(assignment.Task.Uid, out var item)) continue;
                // Existing conflicting assignment dates are rejected by calculation. Newly authored
                // assignments need explicit endpoints so an importer does not choose the project origin.
                if (!assignment.Start.HasValue) assignment.Start = item.Start;
                if (!assignment.Finish.HasValue) assignment.Finish = item.Finish;
            }
        }
        IsScheduleStale = false;
        if (result.CalculatedAssignments) {
            var scheduledAssignments = new HashSet<int>(result.Assignments.Select(a => a.AssignmentUid));
            if (Assignments.All(a => scheduledAssignments.Contains(a.Uid)) && AllTasks.Where(t => t.IsNull != true).All(t => dates.ContainsKey(t.Uid)))
                AreWorkCostTotalsStale = false;
        }
    }
    /// <summary>Calculates and explicitly applies task dates. Throws before mutation when calculation reports errors.</summary>
    public ProjectScheduleResult Recalculate(ProjectScheduleOptions? options = null, CancellationToken cancellationToken = default) {
        var result = CalculateSchedule(options, cancellationToken); ApplySchedule(result, cancellationToken); return result;
    }
}
