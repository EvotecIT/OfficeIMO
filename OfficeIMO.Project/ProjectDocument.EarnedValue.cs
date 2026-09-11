namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Evaluates stored baseline cost curves against independently recorded completion and actual costs. Does not schedule or mutate the project.</summary>
    public ProjectEarnedValueResult AnalyzeEarnedValue(int baselineNumber = 0, DateTime? statusDate = null, int maxCalendarDays = 36600, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("Finish the update scope before analyzing earned value.");
        if (baselineNumber < 0 || baselineNumber > 10) throw new ArgumentOutOfRangeException(nameof(baselineNumber));
        if (maxCalendarDays < 1 || maxCalendarDays > 366000) throw new ArgumentOutOfRangeException(nameof(maxCalendarDays));
        DateTime status = statusDate ?? Settings.StatusDate ?? throw new InvalidOperationException("Earned value requires an explicit StatusDate.");
        ProjectCalendarMath.Local(status); Validate(cancellationToken).ThrowIfErrors();
        long revision = Revision; var results = new List<ProjectTaskEarnedValue>(); var diagnostics = new List<ProjectDiagnostic>();
        var costType = ProjectBaselineTypes.Task(baselineNumber).Cost;
        var assignmentsByTask = Assignments.Where(a => a.Task != null).ToLookup(a => a.Task!.Uid);
        var baselineTimeBases = BaselineCostTimeBases(baselineNumber, cancellationToken);
        foreach (var task in AllTasks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (task.IsNull == true || task.IsActive == false) continue;
            var assignments = assignmentsByTask[task.Uid].ToArray();
            decimal? planned = null, earned = null, actual = null;
            var baseline = task.Baselines.SingleOrDefault(b => (b.Number ?? 0) == baselineNumber);
            void Warn(string message) => diagnostics.Add(new ProjectDiagnostic("PROJECT_EARNED_VALUE_INCOMPLETE", ProjectDiagnosticSeverity.Warning, message, "/Task[UID=" + task.Uid + "]"));
            if (baseline?.Cost == null) { Warn("The selected baseline has no recorded cost."); results.Add(new ProjectTaskEarnedValue(task.Uid, null, null, null, null)); continue; }
            if (task.TimephasedData.Concat(baseline.TimephasedData).Concat(assignments.SelectMany(a => a.TimephasedData)).Any(v => !v.Type.HasValue)) {
                Warn("Timephased records without a type cannot be classified as baseline or actual work and cost.");
                results.Add(new ProjectTaskEarnedValue(task.Uid, baseline.Cost, null, null, null)); continue;
            }
            var calendar = task.Calendar ?? Calendar;
            if (calendar == null) { Warn("Cost-curve integration requires an explicit calendar."); results.Add(new ProjectTaskEarnedValue(task.Uid, baseline.Cost, null, null, null)); continue; }
            var workCalendars = task.IsSummary ? Array.Empty<ProjectCalendar>() : assignments.Where(a => a.Resource?.Type == ProjectResourceType.Work)
                .Select(a => a.Resource!.Calendar ?? Calendar ?? calendar).Distinct().ToArray();
            var effectiveCalendars = task.IgnoreResourceCalendar != true && workCalendars.Length == 1 ?
                (task.Calendar == null ? workCalendars : workCalendars.Concat(new[] { task.Calendar }).ToArray()) : new[] { calendar };
            var math = new ProjectCalendarMath(effectiveCalendars, maxCalendarDays, cancellationToken);
            var curves = task.TimephasedData.Concat(baseline.TimephasedData).Where(v => v.Type == costType).ToArray();
            bool baselineCostCurvesValid = false;
            try {
                if (curves.Length == 0) {
                    if (baseline.Start.HasValue && status <= baseline.Start) planned = 0m;
                    else if (baseline.Finish.HasValue && status >= baseline.Finish) planned = baseline.Cost;
                    else Warn("Baseline cost curves are required to determine the budget at this status boundary.");
                } else {
                    decimal total = curves.Sum(c => ProjectXmlValue.ParseMoney(c.Value ?? throw new InvalidDataException("A baseline cost curve has no value.")));
                    if (Math.Abs(total - baseline.Cost.Value) > .02m) throw new InvalidDataException("Baseline cost curves differ from the stored baseline cost.");
                    baselineCostCurvesValid = true;
                    planned = CostThrough(curves, status, math, baselineTimeBases[task], cancellationToken);
                }
            } catch (Exception exception) when (EarnedValueInputFailure(exception)) { Warn(exception.Message); planned = null; }
            try {
                decimal? completion = CompletionAtStatus(task, assignments, status, Warn);
                if (!completion.HasValue) earned = null;
                else if (task.EarnedValueMethod == ProjectEarnedValueMethod.PhysicalPercentComplete)
                    earned = baseline.Cost * completion.Value / 100m;
                else if (completion == 0) earned = 0m;
                else if (completion == 100) earned = baseline.Cost;
                else if (curves.Length > 0 && baselineCostCurvesValid && baseline.Start.HasValue && baseline.Duration.HasValue) {
                    var duration = baseline.Duration.Value;
                    decimal minutes = duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, this) * completion.Value / 100m;
                    DateTime cutoff = BaselineDurationCutoff(task, baseline, baselineNumber, minutes, math);
                    earned = CostThrough(curves, cutoff, math, baselineTimeBases[task], cancellationToken);
                } else Warn("Duration-based earned value requires baseline start, duration, and cost curves.");
            } catch (Exception exception) when (EarnedValueInputFailure(exception)) { Warn(exception.Message); earned = null; }
            try {
                if (!task.ActualCost.HasValue && (task.IsSummary || task.FixedCost.GetValueOrDefault() != 0m ||
                    (assignments.Length == 0 || task.IsSummary) && (task.ActualStart.HasValue || task.ActualFinish.HasValue || task.ActualWork?.Minutes > 0 || task.ActualDuration?.Value > 0))) {
                    Warn("The task's actual cost contribution is not recorded; assignment costs alone do not establish its actual total.");
                } else if (assignments.Length == 0 || task.IsSummary) {
                    if (ActualCostBoundaryKnown(task.ActualStart, task.Stop, task.ActualFinish, status)) actual = task.ActualCost;
                    else if ((task.ActualCost ?? 0m) == 0) actual = 0m;
                    else Warn("Actual cost at this boundary requires assignment cost curves.");
                } else {
                    decimal sum = 0m, fullAssignmentActual = 0m; bool complete = true;
                    foreach (var assignment in assignments) {
                        cancellationToken.ThrowIfCancellationRequested();
                        var actualCurves = assignment.TimephasedData.Where(v => v.Type == 6).ToArray();
                        fullAssignmentActual += assignment.ActualCost ?? actualCurves.Sum(c => ProjectXmlValue.ParseMoney(c.Value!));
                        if (actualCurves.Length > 0) {
                            if (assignment.ActualCost.HasValue && Math.Abs(actualCurves.Sum(c => ProjectXmlValue.ParseMoney(c.Value!)) - assignment.ActualCost.Value) > .02m)
                                throw new InvalidDataException("Actual cost curves differ from the stored assignment actual cost.");
                            var effective = new List<ProjectCalendar> { task.IgnoreResourceCalendar == true ? calendar : assignment.Resource?.Calendar ?? Calendar ?? calendar };
                            if (task.Calendar != null && task.IgnoreResourceCalendar != true) effective.Add(task.Calendar);
                            sum += CostThrough(actualCurves, status, new ProjectCalendarMath(effective, maxCalendarDays, cancellationToken), task.Duration?.IsElapsed == true, cancellationToken);
                        } else if (!assignment.ActualCost.HasValue && (assignment.ActualStart.HasValue || assignment.ActualFinish.HasValue || assignment.ActualWork?.Minutes > 0 || assignment.PercentWorkComplete > 0))
                            complete = false;
                        else if (ActualCostBoundaryKnown(assignment.ActualStart, assignment.Stop, assignment.ActualFinish, status) || (assignment.ActualCost ?? 0m) == 0m)
                            sum += assignment.ActualCost ?? 0m;
                        else complete = false;
                    }
                    decimal fixedActual = task.ActualCost.HasValue ? task.ActualCost.Value - fullAssignmentActual : 0m;
                    if (fixedActual != 0 && !ActualCostBoundaryKnown(task.ActualStart, task.Stop, task.ActualFinish, status)) complete = false;
                    actual = complete ? sum + fixedActual : (decimal?)null;
                    if (!complete) Warn("Actual cost at this boundary cannot be reconstructed without complete actual cost curves.");
                }
            } catch (Exception exception) when (EarnedValueInputFailure(exception)) {
                Warn(exception.Message); actual = null;
            }
            results.Add(new ProjectTaskEarnedValue(task.Uid, baseline.Cost, planned, earned, actual));
        }
        if (Revision != revision) throw new InvalidOperationException("The document changed during earned-value analysis.");
        return new ProjectEarnedValueResult(revision, status, baselineNumber, results, diagnostics);
    }
    private static bool EarnedValueInputFailure(Exception exception) => exception is InvalidDataException || exception is FormatException ||
        exception is OverflowException || exception is InvalidOperationException || exception is ArgumentException;
    private bool ActualCostBoundaryKnown(DateTime? start, DateTime? stop, DateTime? finish, DateTime status) {
        if (start > status || finish > status || stop > status) return false;
        if (finish <= status) return true;
        return stop <= status && !(Settings.StatusDate > status);
    }
    private decimal? CompletionAtStatus(ProjectTask task, ProjectAssignment[] assignments, DateTime status, Action<string> warn) {
        var actualCurves = assignments.SelectMany(a => a.TimephasedData).Where(v => v.Type == 2).ToArray();
        var starts = assignments.Select(a => a.ActualStart).Concat(actualCurves.Select(v => v.Start)).Append(task.ActualStart)
            .Where(d => d.HasValue).Select(d => d!.Value).ToArray();
        if (starts.Length > 0 && status < starts.Min()) return 0m;
        decimal? completion = task.EarnedValueMethod == ProjectEarnedValueMethod.PhysicalPercentComplete ? task.PhysicalPercentComplete : task.PercentComplete ?? 0;
        if (!completion.HasValue) { warn("Physical earned value requires recorded physical completion."); return null; }
        if (completion == 0 || (completion == 100 && task.ActualFinish <= status)) return completion;
        var boundaries = assignments.SelectMany(a => new[] { a.Stop, a.ActualFinish }).Concat(actualCurves.Select(v => v.Finish))
            .Concat(new[] { Settings.StatusDate, task.Stop, task.ActualFinish });
        if (boundaries.Any(d => d.HasValue && d.Value > status)) {
            warn("Stored completion belongs to a later progress boundary; historical earned value cannot be reconstructed from the current completion percentage.");
            return null;
        }
        return completion;
    }
    private static DateTime BaselineDurationCutoff(ProjectTask task, ProjectBaseline baseline, int number, decimal minutes, ProjectCalendarMath math) {
        var duration = baseline.Duration!.Value;
        if (duration.IsElapsed) return baseline.Start!.Value.Add(ProjectXmlValue.MinutesToSpan(minutes));
        var workCurves = task.TimephasedData.Concat(baseline.TimephasedData)
            .Where(v => v.Type == ProjectBaselineTypes.Task(number).Work).ToArray();
        if (task.IsSummary || workCurves.Length == 0) return math.Add(baseline.Start!.Value, minutes);
        var ranges = new List<ProjectWorkingRange>();
        foreach (var curve in workCurves) {
            if (!curve.Start.HasValue || !curve.Finish.HasValue || curve.Start > curve.Finish || curve.Value == null)
                throw new InvalidDataException("A baseline work curve requires ordered dates and a value.");
            if (ProjectXmlValue.ParseWork(curve.Value).Minutes > 0)
                ranges.Add(new ProjectWorkingRange(curve.Start.Value, curve.Finish.Value));
        }
        ranges = ProjectCalendarMath.Merge(ranges);
        decimal covered = ranges.Sum(r => math.Between(r.Start, r.Finish));
        decimal expected = duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, false, task.Document);
        if (baseline.Finish.HasValue && Math.Abs(math.Between(baseline.Start!.Value, baseline.Finish.Value) - expected) <= .001m)
            return math.Add(baseline.Start.Value, minutes);
        if (Math.Abs(covered - expected) > .001m)
            throw new InvalidDataException("Baseline work curves do not establish the complete baseline duration for earned-value progress.");
        foreach (var range in ranges) {
            decimal span = math.Between(range.Start, range.Finish);
            if (minutes <= span) return math.Add(range.Start, minutes);
            minutes -= span;
        }
        return baseline.Finish ?? baseline.Start!.Value;
    }
    private Dictionary<ProjectTask, bool?> BaselineCostTimeBases(int number, CancellationToken token) {
        var tasks = AllTasks.ToArray(); var bases = new Dictionary<ProjectTask, bool?>();
        foreach (var task in tasks.AsEnumerable().Reverse().Where(t => t.Uid != 0).Concat(tasks.Where(t => t.Uid == 0))) {
            token.ThrowIfCancellationRequested();
            var baseline = task.Baselines.SingleOrDefault(b => b.Number == number);
            bool? elapsed = baseline?.Duration?.IsElapsed ?? (task.Duration?.IsElapsed == true ? (bool?)null : false);
            if (task.IsSummary) {
                var children = task.Uid == 0 ? Tasks.Where(t => t.Uid != 0) : task.Children;
                if (children.Any(child => !bases.TryGetValue(child, out var basis) || basis != elapsed)) elapsed = null;
            }
            bases.Add(task, elapsed);
        }
        return bases;
    }
    private static decimal CostThrough(IEnumerable<ProjectTimephasedValue> curves, DateTime status, ProjectCalendarMath math, bool? elapsed, CancellationToken token) {
        decimal total = 0m;
        foreach (var curve in curves) {
            token.ThrowIfCancellationRequested();
            if (!curve.Start.HasValue || !curve.Finish.HasValue || curve.Finish < curve.Start || curve.Value == null)
                throw new InvalidDataException("A cost curve requires ordered start and finish dates and a value.");
            decimal cost = ProjectXmlValue.ParseMoney(curve.Value);
            if (status < curve.Start) continue;
            if (status >= curve.Finish) { total += cost; continue; }
            if (cost == 0) continue;
            if (!elapsed.HasValue) throw new InvalidDataException("The baseline cost curve's elapsed or working-time basis cannot be established from its owner and rolled-up tasks.");
            decimal span = elapsed.Value ? curve.Finish.Value.Ticks - curve.Start.Value.Ticks : math.Between(curve.Start.Value, curve.Finish.Value);
            if (span <= 0m && cost != 0m) throw new InvalidDataException("A nonzero cost curve has no time in its duration basis.");
            decimal through = elapsed.Value ? status.Ticks - curve.Start.Value.Ticks : math.Between(curve.Start.Value, status);
            if (span > 0m) total += cost * through / span;
        }
        return total;
    }
}
