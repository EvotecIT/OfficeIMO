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
        foreach (var task in AllTasks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (task.IsNull == true || task.IsActive == false) continue;
            decimal? planned = null, earned = null, actual = null;
            var baseline = task.Baselines.SingleOrDefault(b => (b.Number ?? 0) == baselineNumber);
            void Warn(string message) => diagnostics.Add(new ProjectDiagnostic("PROJECT_EARNED_VALUE_INCOMPLETE", ProjectDiagnosticSeverity.Warning, message, "/Task[UID=" + task.Uid + "]"));
            if (baseline?.Cost == null) { Warn("The selected baseline has no recorded cost."); results.Add(new ProjectTaskEarnedValue(task.Uid, null, null, null, null)); continue; }
            var calendar = task.Calendar ?? Calendar;
            if (calendar == null) { Warn("Cost-curve integration requires an explicit calendar."); results.Add(new ProjectTaskEarnedValue(task.Uid, baseline.Cost, null, null, null)); continue; }
            var workCalendars = task.IsSummary ? Array.Empty<ProjectCalendar>() : Assignments.Where(a => a.Task == task && a.Resource?.Type == ProjectResourceType.Work)
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
                    planned = CostThrough(curves, status, math);
                }
            } catch (Exception exception) when (EarnedValueInputFailure(exception)) { Warn(exception.Message); planned = null; }
            try {
                if (task.EarnedValueMethod == ProjectEarnedValueMethod.PhysicalPercentComplete)
                    earned = task.PhysicalPercentComplete.HasValue ? baseline.Cost * task.PhysicalPercentComplete.Value / 100m : null;
                else if ((task.PercentComplete ?? 0) == 0) earned = 0m;
                else if (task.PercentComplete == 100) earned = baseline.Cost;
                else if (curves.Length > 0 && baselineCostCurvesValid && baseline.Start.HasValue && baseline.Duration.HasValue) {
                    var duration = baseline.Duration.Value;
                    decimal minutes = duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, this) * task.PercentComplete!.Value / 100m;
                    DateTime cutoff = BaselineDurationCutoff(task, baseline, baselineNumber, minutes, math);
                    earned = CostThrough(curves, cutoff, math);
                } else Warn("Duration-based earned value requires baseline start, duration, and cost curves.");
            } catch (Exception exception) when (EarnedValueInputFailure(exception)) { Warn(exception.Message); earned = null; }
            try {
                var assignments = Assignments.Where(a => a.Task == task).ToArray();
                if (assignments.Length == 0 || task.IsSummary) {
                    if (task.ActualFinish <= status || task.Stop <= status || task.Finish <= status) actual = task.ActualCost;
                    else if ((task.ActualCost ?? 0m) == 0) actual = 0m;
                    else Warn("Actual cost at this boundary requires assignment cost curves.");
                } else {
                    decimal sum = 0m; bool complete = true;
                    foreach (var assignment in assignments) {
                        cancellationToken.ThrowIfCancellationRequested();
                        var actualCurves = assignment.TimephasedData.Where(v => v.Type == 6).ToArray();
                        if (actualCurves.Length > 0) {
                            var effective = new List<ProjectCalendar> { task.IgnoreResourceCalendar == true ? calendar : assignment.Resource?.Calendar ?? Calendar ?? calendar };
                            if (task.Calendar != null && task.IgnoreResourceCalendar != true) effective.Add(task.Calendar);
                            sum += CostThrough(actualCurves, status, new ProjectCalendarMath(effective, maxCalendarDays, cancellationToken));
                        } else if (assignment.ActualFinish <= status || assignment.Stop <= status || assignment.Finish <= status || (assignment.ActualCost ?? 0m) == 0m)
                            sum += assignment.ActualCost ?? 0m;
                        else complete = false;
                    }
                    decimal fixedActual = (task.ActualCost ?? sum) - assignments.Sum(a => a.ActualCost ?? 0m);
                    if (fixedActual != 0 && !(task.Stop <= status || task.ActualFinish <= status || task.Finish <= status)) complete = false;
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
    private static decimal CostThrough(IEnumerable<ProjectTimephasedValue> curves, DateTime status, ProjectCalendarMath math) {
        decimal total = 0m;
        foreach (var curve in curves) {
            if (!curve.Start.HasValue || !curve.Finish.HasValue || curve.Finish < curve.Start || curve.Value == null)
                throw new InvalidDataException("A cost curve requires ordered start and finish dates and a value.");
            decimal cost = ProjectXmlValue.ParseMoney(curve.Value);
            if (status < curve.Start) continue;
            if (status >= curve.Finish) { total += cost; continue; }
            decimal span = math.Between(curve.Start.Value, curve.Finish.Value);
            if (span <= 0m && cost != 0m) throw new InvalidDataException("A nonzero cost curve has no working time.");
            if (span > 0m) total += cost * math.Between(curve.Start.Value, status) / span;
        }
        return total;
    }
}
