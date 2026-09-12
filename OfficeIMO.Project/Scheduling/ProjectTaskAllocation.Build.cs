namespace OfficeIMO.Project;

internal sealed partial class ProjectTaskAllocation {
    internal bool HasActuals => _entries.Any(e => e.Actual > 0 || e.Assignment.ActualStart.HasValue || e.Assignment.ActualFinish.HasValue)
        || _task.ActualStart.HasValue || _task.ActualFinish.HasValue || _task.ActualDuration?.Value > 0;
    internal Result Build(DateTime anchor, bool forward, bool costs = false) {
        _token.ThrowIfCancellationRequested(); ProjectCalendarMath.Local(anchor);
        if (!forward && HasActuals) throw new NotSupportedException("Backward scheduling of recorded progress requires an explicit remaining-work anchor.");
        if (!forward) {
            // Invert every assignment against the same finish bound, then rebuild all
            // assignments from the latest common task anchor that satisfies that bound.
            var starts = _entries.Where(e => e.Assignment.Resource!.Type == ProjectResourceType.Work)
                .Select(e => e.Calendar.Add(e.RemainingCalendar.Add(anchor, -(e.Curves.Length == 0 ? 0m : e.Curves.Max(c => c.To))), -(e.Assignment.DelayMinutes ?? 0m))).ToArray();
            DateTime start = starts.Length == 0 ? TaskAdd(anchor, -_requestedDuration) : starts.Min();
            if (_task.Type == ProjectTaskType.FixedDuration) start = Min(start, TaskAdd(anchor, -_requestedDuration));
            var inverse = Build(start, true, costs);
            if (inverse.Finish > anchor) throw new InvalidOperationException("Assignment calendars and delays cannot meet the requested finish bound.");
            return inverse;
        }
        var recordedWork = _entries.Where(e => e.Assignment.Resource!.Type == ProjectResourceType.Work)
            .SelectMany(e => e.ActualIntervals).ToArray();
        decimal unrepresentedActualDuration = Math.Max(0m, _actualTaskDuration - UnionMinutes(recordedWork));
        DateTime? actualTaskStart = _task.ActualStart ?? (recordedWork.Length > 0 ? recordedWork.Min(i => i.Start) : (DateTime?)null);
        if (_actualTaskDuration > 0 && !actualTaskStart.HasValue)
            throw new InvalidDataException("Recorded task actual duration requires an actual start or dated actual work before calculation.");
        DateTime? actualTaskEnd = _task.ActualDuration.HasValue && actualTaskStart.HasValue
            ? TaskAdd(actualTaskStart.Value, _actualTaskDuration) : (DateTime?)null;
        if (actualTaskEnd.HasValue && unrepresentedActualDuration > 0 && recordedWork.Length > 0)
            actualTaskEnd = Max(actualTaskEnd.Value, recordedWork.Max(i => i.Finish));
        var plans = new List<ProjectAssignmentSchedule>();
        long intervalCount = 0;
        foreach (var entry in _entries) {
            _token.ThrowIfCancellationRequested();
            var assignment = entry.Assignment;
            decimal delay = assignment.DelayMinutes ?? 0m;
            decimal span = entry.Curves.Length == 0 ? 0m : entry.Curves.Max(c => c.To);
            DateTime origin = forward ? entry.Calendar.Add(anchor, delay) : entry.Calendar.Add(anchor, -span);
            if (entry.ActualIntervals.Length != 0) {
                DateTime actualEnd = entry.ActualIntervals.Max(i => i.Finish);
                origin = Max(origin, actualEnd);
                if (entry.RemainingOrigin.HasValue) origin = Max(origin, entry.RemainingOrigin.Value);
            }
            if (assignment.Resume.HasValue) origin = Max(origin, assignment.Resume.Value);
            // Explicit task progress beyond the recorded assignment intervals anchors the remainder.
            // Inferred actual duration must keep concurrent and discontinuous curves unchanged.
            if (actualTaskEnd.HasValue && unrepresentedActualDuration > 0)
                origin = Max(origin, actualTaskEnd.Value);
            origin = BoundRemainingTaskStart(origin);
            var intervals = new List<ProjectAssignmentInterval>(entry.ActualIntervals);
            decimal regular = entry.Remaining - entry.RemainingOvertime, assignedOvertime = 0m;
            for (int index = 0; index < entry.Curves.Length; index++) {
                var curve = entry.Curves[index];
                decimal overtime = index == entry.Curves.Length - 1 ? entry.RemainingOvertime - assignedOvertime :
                    regular > 0 ? entry.RemainingOvertime * curve.Work / regular : 0;
                DateTime from = entry.RemainingCalendar.Add(origin, curve.From), to = entry.RemainingCalendar.Add(origin, curve.To);
                Expand(entry, entry.RemainingCalendar, from, to, curve.Work + overtime, overtime, false, intervals); assignedOvertime += overtime;
            }
            DateTime start = intervals.Count > 0 ? intervals.Min(i => i.Start) : origin;
            DateTime finish = intervals.Count > 0 ? intervals.Max(i => i.Finish) : entry.RemainingCalendar.Add(origin, span);
            if (assignment.Resource!.Type == ProjectResourceType.Cost) {
                start = anchor;
                finish = anchor;
            } else if ((assignment.ActualStart.HasValue && assignment.ActualStart != start)
                || (assignment.ActualFinish.HasValue && assignment.ActualFinish != finish)) {
                throw new InvalidDataException("Calculated assignment bounds must preserve recorded actual start and finish.");
            }
            var charges = new List<ProjectCostInterval>(); decimal? cost = null, actualCost = null;
            if (costs && assignment.Resource.Type != ProjectResourceType.Cost) (cost, actualCost) = CalculateCosts(entry, intervals, start, finish, charges);
            plans.Add(new ProjectAssignmentSchedule(assignment, start, finish, entry.Units, intervals, charges, cost, actualCost,
                assignment.Resource.Type == ProjectResourceType.Material ? intervals.Sum(i => i.Work.Minutes) / 60m : (decimal?)null, origin));
            intervalCount += (long)intervals.Count + charges.Count;
            CheckCount(intervalCount);
        }
        var workUids = new HashSet<int>(_entries.Where(e => e.Assignment.Resource!.Type == ProjectResourceType.Work).Select(e => e.Assignment.Uid));
        var workIntervals = plans.Where(p => workUids.Contains(p.AssignmentUid)).SelectMany(p => p.Intervals).ToArray();
        decimal actualDuration, remainingDuration, duration;
        DateTime taskStart, taskFinish;
        if (workIntervals.Length == 0) {
            actualDuration = _task.ActualDuration is ProjectDuration actual ? actual.Value * ProjectXmlValue.MinutesPerUnit(actual.Unit, actual.IsElapsed, _document) : 0;
            remainingDuration = RemainingTaskDuration(); duration = actualDuration + remainingDuration;
            if (_task.ActualFinish.HasValue && !_task.ActualStart.HasValue && actualDuration > 0)
                throw new InvalidDataException("A completed task with nonzero actual duration requires an actual start.");
            taskStart = _task.ActualStart ?? _task.ActualFinish ?? anchor;
            DateTime remainingStart = BoundRemainingTaskStart(TaskAdd(taskStart, actualDuration));
            taskFinish = _task.ActualFinish ?? TaskAdd(remainingStart, remainingDuration);
        } else {
            taskStart = _entries.Select((entry, index) => (Entry: entry, Plan: plans[index]))
                .Where(p => workUids.Contains(p.Plan.AssignmentUid)).Min(p => p.Entry.Actual > 0 ? p.Plan.Start :
                    p.Entry.Calendar.Add(p.Plan.Start, -(p.Entry.Assignment.DelayMinutes ?? 0m)));
            taskFinish = workIntervals.Max(i => i.Finish);
            if (_task.ActualStart.HasValue) taskStart = _task.ActualStart.Value;
            actualDuration = UnionMinutes(workIntervals.Where(i => i.IsActual));
            duration = UnionMinutes(workIntervals);
            if (_task.ActualDuration is ProjectDuration recordedDuration)
                actualDuration = recordedDuration.Value * ProjectXmlValue.MinutesPerUnit(recordedDuration.Unit, recordedDuration.IsElapsed, _document);
            duration += unrepresentedActualDuration;
            // Concurrent actual and remaining assignment effort must not count the same task duration twice.
            remainingDuration = Math.Max(0m, duration - actualDuration);
            if (_task.Type == ProjectTaskType.FixedDuration) {
                actualDuration = _task.ActualDuration is ProjectDuration actual ? actual.Value * ProjectXmlValue.MinutesPerUnit(actual.Unit, actual.IsElapsed, _document) : actualDuration;
                remainingDuration = RemainingTaskDuration(); duration = actualDuration + remainingDuration;
                DateTime remainingStart = TaskAdd(taskStart, actualDuration);
                if (recordedWork.Length > 0) remainingStart = Max(remainingStart, recordedWork.Max(i => i.Finish));
                taskFinish = Max(taskFinish, TaskAdd(BoundRemainingTaskStart(remainingStart), remainingDuration));
            }
            if (actualDuration > duration) throw new InvalidDataException("Recorded task actual duration exceeds its calculated duration.");
        }
        if (taskStart > taskFinish) throw new InvalidDataException("Calculated task start must not follow its finish.");
        if (_task.ActualFinish.HasValue && (remainingDuration > 0 || taskFinish != _task.ActualFinish.Value))
            throw new InvalidDataException("A finished task requires completed duration inputs and assignment dates consistent with its actual finish.");
        if (_task.IsManual == true) {
            if (!_task.Start.HasValue || !_task.Finish.HasValue || taskStart < _task.Start || taskFinish > _task.Finish)
                throw new InvalidOperationException("Calculated assignment work cannot fit within the manual task's stored dates.");
            taskStart = _task.Start.Value; taskFinish = _task.Finish.Value;
        }
        for (int index = 0; index < _entries.Length; index++) {
            if (_entries[index].Assignment.Resource!.Type == ProjectResourceType.Work) {
                if (plans[index].Start < taskStart || plans[index].Finish > taskFinish)
                    throw new InvalidDataException("Work assignment dates must fit the final task dates, including recorded actuals.");
                continue;
            }
            if (_entries[index].Assignment.Resource!.Type != ProjectResourceType.Material) continue;
            var material = plans[index];
            if ((workIntervals.Length != 0 && duration != _requestedDuration) || material.Start < taskStart || material.Finish > taskFinish)
                throw new NotSupportedException("Material consumption requires an explicit projection when calculated work changes the declared task duration or material intervals exceed the final task dates.");
        }
        // Cost resources follow the final task span, including calendar, effort, and manual-date adjustments.
        for (int index = 0; index < _entries.Length; index++) {
            var entry = _entries[index];
            if (entry.Assignment.Resource!.Type != ProjectResourceType.Cost) continue;
            _token.ThrowIfCancellationRequested();
            foreach (var date in new[] { entry.Assignment.ActualStart, entry.Assignment.ActualFinish, entry.Assignment.Stop }) {
                if (date.HasValue && (date < taskStart || date > taskFinish))
                    throw new InvalidDataException("Cost assignment actual dates must fit the final task dates.");
            }
            var charges = new List<ProjectCostInterval>(); decimal? cost = null, actualCost = null;
            if (costs) (cost, actualCost) = CalculateCosts(entry, new List<ProjectAssignmentInterval>(), taskStart, taskFinish, charges);
            plans[index] = new ProjectAssignmentSchedule(entry.Assignment, taskStart, taskFinish, entry.Units,
                Array.Empty<ProjectAssignmentInterval>(), charges, cost, actualCost, null);
            intervalCount += charges.Count;
            CheckCount(intervalCount);
        }
        return new Result { Start = taskStart, Finish = taskFinish, Duration = duration, ActualDuration = actualDuration,
            RemainingDuration = remainingDuration, Assignments = plans.ToArray() };
    }
    private static DateTime Max(DateTime first, DateTime second) => first > second ? first : second;
    private static DateTime Min(DateTime first, DateTime second) => first < second ? first : second;
    private DateTime BoundRemainingTaskStart(DateTime start) {
        if (_task.Stop.HasValue) start = Max(start, _task.Stop.Value);
        if (_task.Resume.HasValue) start = Max(start, _task.Resume.Value);
        if (_options.RescheduleRemainingAfterStatusDate) start = Max(start, _document.Settings.StatusDate!.Value);
        return start;
    }
    private DateTime TaskAdd(DateTime date, decimal minutes) => _task.Duration?.IsElapsed == true
        ? date.Add(ProjectXmlValue.MinutesToSpan(minutes)) : _taskCalendar.Add(date, minutes);
    internal ProjectTaskWorkSchedule Totals(Result result) {
        var work = result.Assignments.Where(a => _document.Resources.GetByUid(a.ResourceUid).Type == ProjectResourceType.Work).ToArray();
        decimal totalWork = work.Sum(a => a.Work.Minutes), actualWork = work.Sum(a => a.ActualWork.Minutes);
        if (work.Length == 0) {
            actualWork = _task.ActualWork?.Minutes ?? (_task.Work?.Minutes - _task.RemainingWork?.Minutes) ?? 0m;
            totalWork = _task.Work?.Minutes ?? actualWork + (_task.RemainingWork?.Minutes ?? 0m);
            if (actualWork < 0 || actualWork > totalWork || (_task.RemainingWork.HasValue
                && Math.Abs(totalWork - actualWork - _task.RemainingWork.Value.Minutes) > .001m))
                throw new InvalidDataException("Task total work must equal actual plus remaining work.");
        }
        if ((_task.PercentComplete > 0 && result.ActualDuration == 0 && !_task.ActualFinish.HasValue)
            || (_task.PercentWorkComplete > 0 && actualWork == 0))
            throw new InvalidDataException("Recorded task completion requires explicit actual duration or work before calculation.");
        if (work.Length > 0 && _task.ActualWork.HasValue && Math.Abs(_task.ActualWork.Value.Minutes - actualWork) > .001m)
            throw new InvalidDataException("Assignment actual work must agree with recorded task actual work.");
        if (work.Length > 0 && _task.Type == ProjectTaskType.FixedWork
            && ((_task.Work.HasValue && Math.Abs(_task.Work.Value.Minutes - totalWork) > .001m)
                || (_task.RemainingWork.HasValue && Math.Abs(_task.RemainingWork.Value.Minutes - (totalWork - actualWork)) > .001m)))
            throw new InvalidDataException("Fixed task work must agree with assignment work. Supply complete assignments or explicitly redistribute the declared effort.");
        decimal fixedCost = _task.FixedCost ?? 0m;
        decimal fraction = result.Duration == 0 ? (_task.ActualFinish.HasValue ? 1m : 0m) : result.ActualDuration / result.Duration;
        decimal actualFixed = (_task.FixedCostAccrual ?? ProjectCostAccrual.Prorated) switch {
            ProjectCostAccrual.Start => HasActuals ? fixedCost : 0m,
            ProjectCostAccrual.End => result.RemainingDuration == 0 && HasActuals ? fixedCost : 0m,
            _ => fixedCost * fraction
        };
        decimal? cost = result.Assignments.All(a => a.Cost.HasValue) ? result.Assignments.Sum(a => a.Cost!.Value) + fixedCost : (decimal?)null;
        decimal? actualCost = result.Assignments.All(a => a.ActualCost.HasValue) ? result.Assignments.Sum(a => a.ActualCost!.Value) + actualFixed : (decimal?)null;
        if (!_options.RecalculateActualCosts && _task.ActualCost.HasValue) {
            if (cost.HasValue && actualCost.HasValue) cost += _task.ActualCost.Value - actualCost.Value;
            actualCost = _task.ActualCost;
        }
        return new ProjectTaskWorkSchedule(new ProjectWork(totalWork), new ProjectWork(actualWork), new ProjectWork(totalWork - actualWork),
            result.ActualDuration, result.RemainingDuration, cost, actualCost, _task.PhysicalPercentComplete, _task.Duration?.IsElapsed == true, _task.ActualFinish.HasValue, HasActuals);
    }
    private static decimal UnionMinutes(IEnumerable<ProjectAssignmentInterval> intervals) => ProjectCalendarMath.Merge(intervals
        .Where(i => i.Finish > i.Start && i.Work.Minutes > i.OvertimeWork.Minutes).Select(i => new ProjectWorkingRange(i.Start, i.Finish)).ToList())
        .Sum(i => (i.Finish.Ticks - i.Start.Ticks) / (decimal)TimeSpan.TicksPerMinute);
}
