namespace OfficeIMO.Project;

/// <summary>Canonical assignment-based task arithmetic shared by ordinary scheduling and explicit leveling.</summary>
internal sealed partial class ProjectTaskAllocation {
    private readonly ProjectTask _task;
    private readonly ProjectDocument _document;
    private readonly ProjectScheduleOptions _options;
    private readonly CancellationToken _token;
    private readonly ProjectCalendarMath _taskCalendar;
    private readonly Entry[] _entries;
    private readonly Action<ProjectDiagnostic> _diagnostic;
    private readonly decimal _requestedDuration;
    private sealed class Entry {
        internal ProjectAssignment Assignment = null!;
        internal ProjectCalendarMath Calendar = null!;
        internal ProjectCalendarMath RemainingCalendar = null!;
        internal ProjectResourceTimeline Resource = null!;
        internal decimal Units, Remaining, Actual, RemainingOvertime, ActualOvertime;
        internal DateTime? RemainingOrigin;
        internal Curve[] Curves = Array.Empty<Curve>();
        internal ProjectAssignmentInterval[] ActualIntervals = Array.Empty<ProjectAssignmentInterval>();
        internal bool IsFixedMaterial => Assignment.Resource!.Type == ProjectResourceType.Material && Assignment.HasFixedRateUnits != false;
    }
    private readonly struct Curve {
        internal Curve(decimal from, decimal to, decimal work) { From = from; To = to; Work = work; }
        internal decimal From { get; }
        internal decimal To { get; }
        internal decimal Work { get; }
    }
    internal sealed class Result {
        internal DateTime Start, Finish;
        internal decimal Duration, ActualDuration, RemainingDuration;
        internal ProjectAssignmentSchedule[] Assignments = Array.Empty<ProjectAssignmentSchedule>();
    }
    internal ProjectTaskAllocation(ProjectTask task, ProjectAssignment[] assignments, ProjectScheduleOptions options,
        Func<IEnumerable<ProjectCalendar>, ProjectCalendarMath> calendar, Action<ProjectDiagnostic> diagnostic, CancellationToken token,
        IReadOnlyList<ProjectWorkingRange>? splits = null) {
        _task = task; _document = task.Document; _options = options; _token = token; _diagnostic = diagnostic;
        var projectCalendar = _document.Calendar ?? task.Calendar ?? throw new InvalidOperationException("Assignment calculation requires a project or task calendar.");
        _taskCalendar = calendar(new[] { task.Calendar ?? projectCalendar });
        _requestedDuration = task.Duration is ProjectDuration duration ? duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, _document) : 0m;
        if (task.Duration?.IsElapsed == true && assignments.Any(a => a.Resource?.Type != ProjectResourceType.Cost))
            throw new NotSupportedException("Elapsed tasks with work or material resources require an explicit working-time projection.");
        var work = assignments.Where(a => a.Resource?.Type == ProjectResourceType.Work).ToArray();
        bool redistribute = options.RedistributeEffortDrivenWork && (task.EffortDriven == true || task.Type == ProjectTaskType.FixedWork);
        decimal totalUnits = work.Sum(a => a.Units?.Value ?? a.Resource!.MaxUnits?.Value ?? 1m);
        decimal actual = work.Sum(a => a.ActualWork?.Minutes ?? 0m);
        decimal? remaining = task.RemainingWork?.Minutes ?? (task.Work?.Minutes - actual);
        if (redistribute && (!remaining.HasValue || remaining < 0 || totalUnits <= 0))
            throw new InvalidOperationException("Effort-driven redistribution requires nonnegative stored task work and positive assignment units.");
        _entries = assignments.Select(a => {
            token.ThrowIfCancellationRequested();
            if (a.Resource == null) throw new InvalidOperationException("Resolve every assignment resource before calculating work and costs.");
            var selected = new List<ProjectCalendar>();
            if (task.IgnoreResourceCalendar == true) selected.Add(task.Calendar ?? projectCalendar);
            else {
                selected.Add(a.Resource.Calendar ?? projectCalendar);
                if (task.Calendar != null) selected.Add(task.Calendar);
            }
            var entry = new Entry { Assignment = a, Calendar = calendar(selected), Resource = new ProjectResourceTimeline(a.Resource),
                Units = a.Units?.Value ?? a.Resource.MaxUnits?.Value ?? 1m };
            Prepare(entry, redistribute && a.Resource.Type == ProjectResourceType.Work ? remaining!.Value * entry.Units / totalUnits : (decimal?)null);
            entry.RemainingCalendar = splits == null || splits.Count == 0 ? entry.Calendar : entry.Calendar.Excluding(splits);
            return entry;
        }).ToArray();
    }
    private void Warn(string code, string message, ProjectAssignment assignment) => _diagnostic(new ProjectDiagnostic(code,
        ProjectDiagnosticSeverity.Warning, message, "/Assignment[UID=" + assignment.Uid + "]"));
    private void Prepare(Entry entry, decimal? redistributed) {
        var assignment = entry.Assignment; var resource = assignment.Resource!;
        entry.Actual = assignment.ActualWork?.Minutes ?? 0m;
        if (assignment.PercentWorkComplete > 0 && entry.Actual <= 0)
            throw new InvalidDataException("Recorded assignment completion requires explicit actual work before calculation.");
        entry.ActualOvertime = assignment.ActualOvertimeWork?.Minutes ?? 0m;
        entry.RemainingOvertime = (assignment.OvertimeWork?.Minutes ?? entry.ActualOvertime) - entry.ActualOvertime;
        if (entry.ActualOvertime > entry.Actual || entry.RemainingOvertime < 0) throw new InvalidDataException("Overtime must be part of total work and actual overtime must be part of actual work.");
        if (resource.Type == ProjectResourceType.Cost) return;
        if (assignment.Work.HasValue && assignment.RemainingWork.HasValue
            && Math.Abs(assignment.Work.Value.Minutes - entry.Actual - assignment.RemainingWork.Value.Minutes) > .001m)
            throw new InvalidDataException("Assignment total work must equal actual plus remaining work.");
        decimal? stored = assignment.RemainingWork?.Minutes ?? (assignment.Work?.Minutes - entry.Actual);
        if (resource.Type == ProjectResourceType.Material) {
            decimal quantity = assignment.HasFixedRateUnits == false ? VariableQuantity(assignment, _requestedDuration) : entry.Units;
            entry.Remaining = stored ?? Math.Max(0, quantity * 60m - entry.Actual);
        } else entry.Remaining = redistributed ?? stored ?? (_requestedDuration * entry.Units - entry.Actual);
        if (entry.Remaining < 0 || entry.RemainingOvertime > entry.Remaining) throw new InvalidDataException("Remaining work and overtime are inconsistent.");
        if (assignment.ActualFinish.HasValue && entry.Remaining > 0) throw new InvalidDataException("A completed assignment cannot have remaining work.");
        if (resource.Type == ProjectResourceType.Work && entry.Units <= 0 && entry.Remaining > entry.RemainingOvertime)
            throw new InvalidOperationException("Positive remaining regular work requires positive assignment units.");
        entry.ActualIntervals = ActualIntervals(entry);
        var storedCurves = assignment.TimephasedData.Where(v => v.Type == 1).OrderBy(v => v.Start).ToArray();
        if (storedCurves.Length != 0 && !redistributed.HasValue) {
            entry.RemainingOrigin = storedCurves[0].Start ?? throw new InvalidDataException("Timephased work requires start and finish dates.");
            var curves = new List<Curve>(); DateTime? last = null;
            foreach (var item in storedCurves) {
                _token.ThrowIfCancellationRequested();
                RequireInterval(item); if (last > item.Start) throw new InvalidDataException("Remaining-work intervals overlap."); last = item.Finish;
                curves.Add(new Curve(entry.Calendar.Between(entry.RemainingOrigin.Value, item.Start!.Value),
                    entry.Calendar.Between(entry.RemainingOrigin.Value, item.Finish!.Value), WorkValue(item)));
            }
            decimal regular = entry.Remaining - entry.RemainingOvertime;
            if (Math.Abs(curves.Sum(c => c.Work) - regular) > 0.001m) throw new InvalidDataException("Timephased remaining regular work differs from the stored remaining work minus overtime.");
            entry.Curves = curves.ToArray();
        } else {
            decimal regular = entry.Remaining - entry.RemainingOvertime;
            decimal duration = resource.Type == ProjectResourceType.Material ? _requestedDuration :
                _task.Type == ProjectTaskType.FixedDuration ? RemainingTaskDuration() : entry.Units > 0 ? regular / entry.Units : 0;
            if (_task.Type == ProjectTaskType.FixedDuration && resource.Type == ProjectResourceType.Work && duration > 0 && redistributed.HasValue)
                entry.Units = regular / duration;
            entry.Curves = entry.Remaining == 0 ? Array.Empty<Curve>() : entry.IsFixedMaterial && duration == 0
                ? new[] { new Curve(0, 0, regular) }
                : NamedCurves(assignment.WorkContour ?? ProjectWorkContour.Flat, duration, regular);
        }
    }
    private decimal RemainingTaskDuration() {
        if (_task.RemainingDuration is ProjectDuration remaining) return remaining.Value * ProjectXmlValue.MinutesPerUnit(remaining.Unit, remaining.IsElapsed, _document);
        decimal actual = _task.ActualDuration is ProjectDuration duration ? duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, _document) : 0m;
        return Math.Max(0, _requestedDuration - actual);
    }
    private decimal VariableQuantity(ProjectAssignment assignment, decimal minutes) {
        int scale = assignment.MaterialRateScale ?? throw new InvalidOperationException("Variable material consumption requires an explicit rate scale.");
        decimal perUnit = scale switch { 1 => 1m, 2 => 60m, 3 => _document.Settings.MinutesPerDay ?? 480,
            4 => _document.Settings.MinutesPerWeek ?? 2400, 5 => (_document.Settings.MinutesPerDay ?? 480) * (_document.Settings.DaysPerMonth ?? 20),
            _ => throw new NotSupportedException("Unsupported variable-material rate scale.") };
        return minutes / perUnit * (assignment.Units?.Value ?? 0m);
    }
    private static void RequireInterval(ProjectTimephasedValue value) {
        if (!value.Start.HasValue || !value.Finish.HasValue || value.Start > value.Finish || value.Value == null)
            throw new InvalidDataException("Timephased work requires an ordered range and a value.");
        ProjectCalendarMath.Local(value.Start.Value); ProjectCalendarMath.Local(value.Finish.Value);
    }
    private static decimal WorkValue(ProjectTimephasedValue value) {
        if (value.Value == null) throw new InvalidDataException("A work interval has no value.");
        return ProjectXmlValue.ParseWork(value.Value).Minutes;
    }
}
