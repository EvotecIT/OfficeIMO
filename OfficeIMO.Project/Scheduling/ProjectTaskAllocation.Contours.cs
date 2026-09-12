namespace OfficeIMO.Project;

internal sealed partial class ProjectTaskAllocation {
    private Curve[] NamedCurves(ProjectWorkContour contour, decimal duration, decimal regularWork) {
        if (contour == ProjectWorkContour.Custom) throw new InvalidOperationException("A custom contour requires explicit remaining-work timephased intervals.");
        if (contour == ProjectWorkContour.Turtle) throw new NotSupportedException("Turtle contour calculation requires explicit timephased intervals to retain producer rounding.");
        decimal[] weights = contour switch {
            ProjectWorkContour.Flat => new[] { 1m },
            ProjectWorkContour.BackLoaded => new[] { .1m, .15m, .25m, .5m, .5m, .75m, .75m, 1m, 1m, 1m },
            ProjectWorkContour.FrontLoaded => new[] { 1m, 1m, 1m, .75m, .75m, .5m, .5m, .25m, .15m, .1m },
            ProjectWorkContour.DoublePeak => new[] { .25m, .5m, 1m, .5m, .25m, .25m, .5m, 1m, .5m, .25m },
            ProjectWorkContour.EarlyPeak => new[] { .25m, .5m, 1m, 1m, .75m, .5m, .5m, .25m, .15m, .1m },
            ProjectWorkContour.LatePeak => new[] { .1m, .15m, .25m, .5m, .5m, .75m, 1m, 1m, .5m, .25m },
            ProjectWorkContour.Bell => new[] { .1m, .2m, .4m, .8m, 1m, 1m, .8m, .4m, .2m, .1m },
            _ => throw new NotSupportedException("Unknown work contour.")
        };
        decimal sum = weights.Sum();
        if (_task.Type != ProjectTaskType.FixedDuration) duration *= weights.Length / sum;
        var result = new Curve[weights.Length]; decimal used = 0;
        for (int index = 0; index < result.Length; index++) {
            decimal amount = index == result.Length - 1 ? regularWork - used : regularWork * weights[index] / sum;
            result[index] = new Curve(duration * index / result.Length, duration * (index + 1) / result.Length, amount); used += amount;
        }
        return result;
    }
    private ProjectAssignmentInterval[] ActualIntervals(Entry entry) {
        var assignment = entry.Assignment;
        if (entry.Actual == 0 && assignment.Stop.HasValue)
            throw new NotSupportedException("A stopped work or material assignment requires explicit nonzero actual work before calculation.");
        var actual = assignment.TimephasedData.Where(v => v.Type == 2).OrderBy(v => v.Start).ToArray();
        var overtime = assignment.TimephasedData.Where(v => v.Type == 3).OrderBy(v => v.Start).ToArray();
        ValidateActualCurves(entry, actual, overtime);
        var result = new List<ProjectAssignmentInterval>();
        if (actual.Length == 0) {
            if (overtime.Length != 0) throw new InvalidDataException("Timephased actual overtime requires explicit actual-work intervals before calculation.");
            if (entry.Actual == 0) return result.ToArray();
            DateTime start = assignment.ActualStart ?? _task.ActualStart ?? throw new InvalidOperationException("Actual work requires an explicit actual start or timephased actual work.");
            decimal regular = entry.Actual - entry.ActualOvertime;
            if (regular > 0 && entry.Units <= 0) throw new InvalidOperationException("Recorded actual work needs positive units or explicit intervals.");
            DateTime finish = assignment.Stop ?? assignment.ActualFinish ?? entry.Calendar.Add(start, entry.Units > 0 ? regular / entry.Units : 0m);
            Expand(entry, entry.Calendar, start, finish, entry.Actual, entry.ActualOvertime, true, result);
            return ValidateActualBounds(assignment, result);
        }
        DateTime? previous = null;
        foreach (var item in actual) {
            _token.ThrowIfCancellationRequested(); RequireInterval(item);
            if (previous > item.Start) throw new InvalidDataException("Actual-work intervals overlap."); previous = item.Finish;
            var boundaries = new SortedSet<DateTime> { item.Start!.Value, item.Finish!.Value };
            foreach (var over in overtime) {
                RequireInterval(over);
                if (over.Start > item.Start && over.Start < item.Finish) boundaries.Add(over.Start.Value);
                if (over.Finish > item.Start && over.Finish < item.Finish) boundaries.Add(over.Finish.Value);
            }
            decimal duration = entry.Calendar.Between(item.Start.Value, item.Finish.Value), amount = WorkValue(item);
            var points = boundaries.ToArray();
            if (points.Length == 1) { Expand(entry, entry.Calendar, points[0], points[0], amount, entry.IsFixedMaterial ? 0 : amount, true, result); continue; }
            for (int i = 1; i < points.Length; i++) {
                decimal work = duration > 0 ? amount * entry.Calendar.Between(points[i - 1], points[i]) / duration : 0m, overAmount = 0m;
                foreach (var over in overtime.Where(o => o.Start <= points[i - 1] && o.Finish >= points[i])) {
                    decimal span = entry.Calendar.Between(over.Start!.Value, over.Finish!.Value);
                    if (span > 0) overAmount += WorkValue(over) * entry.Calendar.Between(points[i - 1], points[i]) / span;
                }
                Expand(entry, entry.Calendar, points[i - 1], points[i], work, overAmount, true, result);
            }
        }
        if (Math.Abs(result.Sum(i => i.Work.Minutes) - entry.Actual) > .001m || Math.Abs(result.Sum(i => i.OvertimeWork.Minutes) - entry.ActualOvertime) > .001m)
            throw new InvalidDataException("Timephased actual work differs from stored actual work or overtime.");
        return ValidateActualBounds(assignment, result);
    }
    private void ValidateActualCurves(Entry entry, ProjectTimephasedValue[] actual, ProjectTimephasedValue[] overtime) {
        foreach (var curves in new[] { actual, overtime }) {
            DateTime? previous = null;
            foreach (var curve in curves) {
                _token.ThrowIfCancellationRequested(); RequireInterval(curve);
                decimal amount = WorkValue(curve);
                if (amount < 0 || (previous.HasValue && previous > curve.Start))
                    throw new InvalidDataException("Actual-work and overtime curves require nonnegative, nonoverlapping intervals.");
                previous = curve.Finish;
                if (amount > 0 && curve.Start != curve.Finish && entry.Calendar.Between(curve.Start!.Value, curve.Finish!.Value) <= 0)
                    throw new InvalidDataException("Nonzero actual work falls outside the effective working calendar.");
            }
        }
        if ((actual.Length > 0 && Math.Abs(actual.Sum(WorkValue) - entry.Actual) > .001m)
            || (overtime.Length > 0 && Math.Abs(overtime.Sum(WorkValue) - entry.ActualOvertime) > .001m))
            throw new InvalidDataException("Timephased actual work differs from stored actual work or overtime.");
        foreach (var over in overtime) {
            _token.ThrowIfCancellationRequested();
            if (WorkValue(over) == 0) continue;
            DateTime start = over.Start!.Value, finish = over.Finish!.Value;
            bool covered;
            if (start == finish) covered = actual.Any(a => a.Start == start && a.Finish == finish);
            else {
                decimal overlap = 0;
                foreach (var item in actual) {
                    DateTime from = Max(start, item.Start!.Value), to = Min(finish, item.Finish!.Value);
                    if (from < to) overlap += entry.Calendar.Between(from, to);
                }
                covered = overlap == entry.Calendar.Between(start, finish);
            }
            if (!covered) throw new InvalidDataException("Every nonzero actual-overtime curve requires complete actual-work interval coverage.");
        }
    }
    private static ProjectAssignmentInterval[] ValidateActualBounds(ProjectAssignment assignment, List<ProjectAssignmentInterval> intervals) {
        if (intervals.Count > 0) {
            DateTime start = intervals.Min(i => i.Start), finish = intervals.Max(i => i.Finish);
            if ((assignment.ActualStart.HasValue && assignment.ActualStart != start)
                || (assignment.ActualFinish.HasValue && assignment.ActualFinish != finish)
                || (assignment.Stop.HasValue && assignment.Stop != finish))
                throw new InvalidDataException("Actual-work intervals must agree with recorded assignment actual start, actual finish, and stop dates.");
        }
        return intervals.ToArray();
    }
    private void Expand(Entry entry, ProjectCalendarMath calendar, DateTime start, DateTime finish, decimal work, decimal overtime, bool actual, List<ProjectAssignmentInterval> output) {
        if (work == 0 && overtime == 0) return;
        if (overtime > work || work < 0 || overtime < 0) throw new InvalidDataException("Interval work and overtime are inconsistent.");
        if (start == finish) {
            if (work != overtime && !entry.IsFixedMaterial) throw new InvalidDataException("Nonzero regular work requires a nonempty working interval.");
            output.Add(new ProjectAssignmentInterval(start, finish, work, overtime, actual)); CheckCount(output.Count); return;
        }
        decimal total = calendar.Between(start, finish);
        if (total <= 0) throw new InvalidDataException("Nonzero work falls outside the effective working calendar.");
        decimal allocated = 0, allocatedOvertime = 0;
        var ranges = new List<ProjectWorkingRange>();
        for (DateTime day = start.Date; day <= finish.Date; day = day.AddDays(1)) {
            _token.ThrowIfCancellationRequested();
            foreach (var range in calendar.Day(day)) {
                DateTime from = range.Start > start ? range.Start : start, to = range.Finish < finish ? range.Finish : finish;
                if (to > from) { ranges.Add(new ProjectWorkingRange(from, to)); CheckCount(ranges.Count); }
            }
        }
        for (int index = 0; index < ranges.Count; index++) {
            var range = ranges[index]; decimal fraction = (range.Finish.Ticks - range.Start.Ticks) / (decimal)TimeSpan.TicksPerMinute / total;
            decimal part = index == ranges.Count - 1 ? work - allocated : work * fraction;
            decimal over = index == ranges.Count - 1 ? overtime - allocatedOvertime : overtime * fraction;
            output.Add(new ProjectAssignmentInterval(range.Start, range.Finish, part, over, actual)); CheckCount(output.Count);
            allocated += part; allocatedOvertime += over;
        }
    }
    private void CheckCount(long count) { if (count > _options.MaxIntervals) throw new InvalidOperationException("Assignment calculation exceeds MaxIntervals."); }
}
