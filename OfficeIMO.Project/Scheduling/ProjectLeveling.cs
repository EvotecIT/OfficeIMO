using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Bounds and ordering for explicit resource leveling. Ordinary schedule calculation never invokes leveling.</summary>
public sealed class ProjectLevelingOptions {
    /// <summary>Maximum recalculation attempts; exhaustion returns an error without changing the document.</summary>
    public int MaxIterations { get; set; } = 1000;
    /// <summary>Maximum calendar days that any task's remaining-work anchor may move from its original proposed start.</summary>
    public int MaxDelayDays { get; set; } = 366;
    /// <summary>Keep each moved task within its original late-finish bound.</summary>
    public bool WithinAvailableSlack { get; set; }
    /// <summary>Use task priority before UID when choosing which task to delay. Priority 1000 always prevents movement.</summary>
    public bool UseTaskPriority { get; set; } = true;
    /// <summary>Try interrupting remaining work at a conflict before delaying the whole remaining-work anchor. Applies to tasks with work-resource assignments only.</summary>
    public bool AllowSplitting { get; set; }
    /// <summary>Maximum generated interruptions across the proposal.</summary>
    public int MaxSplits { get; set; } = 1000;
    /// <summary>Calculation settings; independent assignment calculation is required.</summary>
    public ProjectScheduleOptions ScheduleOptions { get; set; } = new ProjectScheduleOptions { CalculateAssignments = true };
}

/// <summary>An interruption proposed for remaining work; recorded actual work is unchanged.</summary>
public sealed class ProjectLevelingSplit {
    internal ProjectLevelingSplit(int taskUid, DateTime start, DateTime finish) { TaskUid = taskUid; Start = start; Finish = finish; }
    /// <summary>Stable task identity.</summary>
    public int TaskUid { get; }
    /// <summary>Inclusive beginning of the interruption.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive end of the interruption.</summary>
    public DateTime Finish { get; }
}

/// <summary>One task's proposed remaining-work anchor introduced by leveling.</summary>
public sealed class ProjectLevelingMove {
    internal ProjectLevelingMove(int uid, DateTime original, DateTime anchor) { TaskUid = uid; OriginalStart = original; NotBefore = anchor; }
    /// <summary>Stable task identity.</summary>
    public int TaskUid { get; }
    /// <summary>Start in the initial schedule proposal, before leveling.</summary>
    public DateTime OriginalStart { get; }
    /// <summary>Earliest permitted remaining-work anchor in this proposal.</summary>
    public DateTime NotBefore { get; }
}

/// <summary>Immutable, revision-bound leveling proposal. Apply Schedule explicitly only when Report contains no errors.</summary>
public sealed class ProjectLevelingResult {
    internal ProjectLevelingResult(ProjectScheduleResult schedule, ProjectResourceAllocationResult? capacity, IEnumerable<ProjectLevelingMove> moves, IEnumerable<ProjectDiagnostic> diagnostics,
        IEnumerable<ProjectLevelingSplit>? splits = null) {
        Schedule = schedule; Capacity = capacity; Moves = new ReadOnlyCollection<ProjectLevelingMove>(moves.ToArray());
        Splits = new ReadOnlyCollection<ProjectLevelingSplit>((splits ?? Array.Empty<ProjectLevelingSplit>()).ToArray());
        Report = new ProjectReport(schedule.ModelRevision, schedule.Report.Diagnostics.Concat(diagnostics));
    }
    /// <summary>Last proposed schedule; use ApplyLeveling to enforce the leveling report as well as the schedule report.</summary>
    public ProjectScheduleResult Schedule { get; }
    /// <summary>Final regular-work capacity analysis, when schedule calculation succeeded.</summary>
    public ProjectResourceAllocationResult? Capacity { get; }
    /// <summary>Moves in task UID order; actual work is never moved.</summary>
    public IReadOnlyList<ProjectLevelingMove> Moves { get; }
    /// <summary>Generated remaining-work interruptions, persisted as assignment timephased work when applied.</summary>
    public IReadOnlyList<ProjectLevelingSplit> Splits { get; }
    /// <summary>Calculation conflicts, remaining overloads, or exhausted limits.</summary>
    public ProjectReport Report { get; }
}

internal static class ProjectLeveler {
    internal static ProjectLevelingResult Calculate(ProjectDocument document, ProjectLevelingOptions options, CancellationToken token) {
        var initial = document.CalculateSchedule(options.ScheduleOptions, token);
        var current = initial; ProjectResourceAllocationResult? capacity = null;
        var anchors = new Dictionary<int, DateTime>();
        var splits = new Dictionary<int, IReadOnlyList<ProjectWorkingRange>>();
        var originals = initial.Tasks.ToDictionary(t => t.TaskUid);
        bool WithinDelayLimits(ProjectScheduleResult proposal) => proposal.Tasks.Where(t => !t.IsSummary).All(t => {
            var original = originals[t.TaskUid];
            return !(t.Start > original.Start && (t.Start - original.Start).TotalDays > options.MaxDelayDays)
                && !(t.ForwardAnchor > original.ForwardAnchor && (t.ForwardAnchor - original.Start).TotalDays > options.MaxDelayDays);
        });
        ProjectLevelingResult Result(string? error = null) => new ProjectLevelingResult(current, capacity,
            anchors.OrderBy(p => p.Key).Select(p => new ProjectLevelingMove(p.Key, originals[p.Key].Start, p.Value)),
            error == null ? Array.Empty<ProjectDiagnostic>() : new[] { new ProjectDiagnostic("PROJECT_LEVELING_INCOMPLETE", ProjectDiagnosticSeverity.Error, error, "/Project") },
            splits.OrderBy(p => p.Key).SelectMany(p => p.Value.Select(r => new ProjectLevelingSplit(p.Key, r.Start, r.Finish))));
        if (current.Report.HasErrors) return Result();
        if (document.Settings.ScheduleFromStart == false) return Result("Leveling requires forward scheduling.");
        for (int iteration = 0; iteration <= options.MaxIterations; iteration++) {
            token.ThrowIfCancellationRequested();
            capacity = document.AnalyzeResourceAllocation(current, options.ScheduleOptions.MaxIntervals, token);
            var conflict = capacity.Overallocations.OrderBy(i => i.Start).ThenBy(i => i.ResourceUid).FirstOrDefault();
            if (conflict == null) return Result();
            if (document.Resources.GetByUid(conflict.ResourceUid).CanLevel == false)
                return Result("The overallocated resource explicitly prohibits leveling.");
            if (iteration == options.MaxIterations) return Result("The iteration limit was reached before every resource conflict was resolved.");
            var involved = new HashSet<int>(conflict.AssignmentUids);
            var candidates = current.Assignments.Where(a => involved.Contains(a.AssignmentUid) && a.Intervals.Any(i => !i.IsActual && i.Start < conflict.Finish && i.Finish > conflict.Start))
                .Select(a => document.Tasks.GetByUid(a.TaskUid)).Distinct()
                .Where(t => t.IsManual != true && (t.Priority ?? 500) < 1000 && t.ConstraintType != ProjectConstraintType.AsLateAsPossible)
                .OrderBy(t => options.UseTaskPriority ? t.Priority ?? 500 : 0).ThenByDescending(t => t.Uid).ToArray();
            if (candidates.Length == 0) return Result("The resource conflict involves only recorded actuals, manual tasks, late-scheduled tasks, or tasks with priority 1000.");
            bool moved = false;
            foreach (var task in candidates) {
                var taskAssignments = current.Assignments.Where(a => a.TaskUid == task.Uid).ToArray();
                if (options.AllowSplitting && task.LevelingCanSplit != false &&
                    taskAssignments.All(a => document.Resources.GetByUid(a.ResourceUid).Type == ProjectResourceType.Work) &&
                    taskAssignments.Any(a => a.Intervals.Any(i => !i.IsActual && i.Start < conflict.Start && i.Work.Minutes > 0)) &&
                    splits.Sum(p => p.Value.Count) < options.MaxSplits) {
                    var ranges = splits.TryGetValue(task.Uid, out var prior) ? prior.ToList() : new List<ProjectWorkingRange>();
                    ranges.Add(new ProjectWorkingRange(conflict.Start, conflict.Finish));
                    var nextSplits = new Dictionary<int, IReadOnlyList<ProjectWorkingRange>>(splits) { [task.Uid] = ProjectCalendarMath.Merge(ranges) };
                    var splitSchedule = new ProjectScheduler(document, options.ScheduleOptions, token, anchors, splits: nextSplits).Calculate();
                    if (!splitSchedule.Report.HasErrors && WithinDelayLimits(splitSchedule) &&
                        splitSchedule.Tasks.Single(t => t.TaskUid == task.Uid).Finish <= originals[task.Uid].Finish.AddDays(options.MaxDelayDays) &&
                        (!options.WithinAvailableSlack || splitSchedule.Tasks.All(t => t.IsSummary || t.Finish <= originals[t.TaskUid].LateFinish))) {
                        current = splitSchedule; splits = nextSplits; moved = true; break;
                    }
                }
                DateTime anchor = conflict.Finish;
                if (anchors.TryGetValue(task.Uid, out var previous) && anchor <= previous) continue;
                if ((anchor - originals[task.Uid].Start).TotalDays > options.MaxDelayDays) continue;
                var nextAnchors = new Dictionary<int, DateTime>(anchors) { [task.Uid] = anchor };
                var next = new ProjectScheduler(document, options.ScheduleOptions, token, nextAnchors, splits: splits).Calculate();
                if (next.Report.HasErrors || !WithinDelayLimits(next)) continue;
                if (options.WithinAvailableSlack && next.Tasks.Any(t => !t.IsSummary && t.Finish > originals[t.TaskUid].LateFinish)) continue;
                current = next; anchors = nextAnchors; moved = true; break;
            }
            if (!moved) return Result("No permitted task move resolves the current conflict within constraints, delay limits, and the selected slack policy.");
        }
        throw new InvalidOperationException("Unreachable leveling state.");
    }
}
