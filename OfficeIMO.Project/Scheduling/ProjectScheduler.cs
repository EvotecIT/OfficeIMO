namespace OfficeIMO.Project;

internal sealed partial class ProjectScheduler {
    private readonly ProjectDocument _document;
    private readonly ProjectScheduleOptions _options;
    private readonly CancellationToken _token;
    private readonly long _revision;
    private readonly List<ProjectDiagnostic> _diagnostics = new List<ProjectDiagnostic>();
    private readonly Dictionary<ProjectTask, Node> _nodes = new Dictionary<ProjectTask, Node>();
    private readonly List<Node> _order = new List<Node>();
    private readonly Dictionary<string, ProjectCalendarMath> _calendars = new Dictionary<string, ProjectCalendarMath>(StringComparer.Ordinal);
    private sealed class Node {
        internal ProjectTask Task = null!;
        internal ProjectCalendarMath Calendar = null!;
        internal decimal Minutes;
        internal bool Elapsed;
        internal DateTime EarlyStart, EarlyFinish, LateStart, LateFinish, Start, Finish;
        internal readonly List<ProjectDependency> In = new List<ProjectDependency>(), Out = new List<ProjectDependency>();
    }
    internal ProjectScheduler(ProjectDocument document, ProjectScheduleOptions options, CancellationToken token) {
        _document = document; _options = new ProjectScheduleOptions { MaxCalendarDays = options.MaxCalendarDays, CriticalSlackMinutes = options.CriticalSlackMinutes };
        _token = token; _revision = document.Revision;
    }
    private void Error(string code, string message, ProjectTask? task = null) => _diagnostics.Add(new ProjectDiagnostic(code,
        ProjectDiagnosticSeverity.Error, message, task == null ? "/Project" : "/Task[UID=" + task.Uid + "]"));
    private bool Failed => _diagnostics.Any(d => d.Severity == ProjectDiagnosticSeverity.Error);
    internal ProjectScheduleResult Calculate() {
        _token.ThrowIfCancellationRequested();
        _diagnostics.AddRange(_document.Validate(_token).Diagnostics.Where(d => d.Severity == ProjectDiagnosticSeverity.Error));
        if (Failed) return Result();
        try {
            Prepare();
            if (Failed || _order.Count == 0) return Result();
            bool forward = _document.Settings.ScheduleFromStart ?? true;
            DateTime origin, horizon;
            if (forward) {
                origin = _document.Settings.StartDate ?? throw new InvalidOperationException("An explicit project StartDate is required.");
                Forward(origin);
                horizon = _order.Max(n => n.EarlyFinish);
                Backward(horizon);
            } else {
                horizon = _document.Settings.FinishDate ?? throw new InvalidOperationException("An explicit project FinishDate is required.");
                Backward(horizon);
                origin = _order.Min(n => n.LateStart);
                Forward(origin);
            }
            foreach (var node in _order) {
                bool late = !forward || node.Task.ConstraintType == ProjectConstraintType.AsLateAsPossible;
                node.Start = late ? node.LateStart : node.EarlyStart;
                node.Finish = late ? node.LateFinish : node.EarlyFinish;
            }
            CheckFinalBounds();
            return Result(BuildResults(horizon));
        } catch (Exception exception) when (exception is InvalidDataException || exception is InvalidOperationException ||
            exception is NotSupportedException || exception is ArgumentException || exception is OverflowException) {
            Error("PROJECT_CALCULATION_UNSUPPORTED", exception.Message); return Result();
        }
    }
    private ProjectScheduleResult Result(IEnumerable<ProjectTaskSchedule>? tasks = null) {
        _token.ThrowIfCancellationRequested();
        if (_document.Revision != _revision) throw new InvalidOperationException("The document changed during schedule calculation.");
        var calculated = tasks?.ToArray() ?? Array.Empty<ProjectTaskSchedule>();
        CheckStoredAssignmentDates(calculated);
        return new ProjectScheduleResult(_document, _revision, calculated, _diagnostics);
    }
    private void Prepare() {
        var all = _document.AllTasks.ToArray();
        CheckSourceProfile();
        var assignments = _document.Assignments.Where(a => a.Task != null && a.Resource?.Type == ProjectResourceType.Work)
            .GroupBy(a => a.Task!).ToDictionary(g => g.Key, g => g.ToArray());
        foreach (var task in all) {
            _token.ThrowIfCancellationRequested();
            if (task.IsSummary || task.IsNull == true || task.IsActive == false || task.Uid == 0) continue;
            var calendar = task.Calendar ?? _document.Calendar;
            if (calendar == null) { Error("PROJECT_CALCULATION_CALENDAR", "Set an explicit task or project calendar.", task); continue; }
            if (task.IsManual != true && !task.Duration.HasValue) { Error("PROJECT_CALCULATION_DURATION", "Automatic tasks require an explicit duration.", task); continue; }
            if (task.IsManual == true && (!task.Start.HasValue || !task.Finish.HasValue)) { Error("PROJECT_MANUAL_DATES", "Manual tasks require both stored dates.", task); continue; }
            if (task.ActualStart.HasValue || task.ActualFinish.HasValue || task.ActualWork?.Minutes > 0 || task.ActualDuration?.Value > 0 || task.PercentComplete > 0 || task.PercentWorkComplete > 0)
                Error("PROJECT_PROGRESS_SCHEDULING", "Progress and status-date rescheduling are not qualified by the date-only scheduling profile.", task);
            var calendars = new List<ProjectCalendar> { calendar };
            assignments.TryGetValue(task, out var assigned);
            if (assigned != null) {
                var resourceCalendars = assigned.Select(a => a.Resource!.Calendar ?? _document.Calendar ?? calendar).Distinct().ToArray();
                if (resourceCalendars.Length > 1) Error("PROJECT_RESOURCE_CALENDAR_SCHEDULING", "Assignments with different resource calendars require independent assignment scheduling.", task);
                if (task.Calendar == null) calendars.Clear();
                calendars.AddRange(resourceCalendars);
            }
            var math = CalendarMath(calendars);
            var duration = task.Duration ?? ProjectDuration.WorkingMinutes(math.Between(task.Start!.Value, task.Finish!.Value));
            if (task.Type == ProjectTaskType.FixedWork && task.IsManual != true) {
                decimal units = assigned?.Sum(a => a.Units?.Value ?? 0m) ?? 0m;
                if (!task.Work.HasValue || units <= 0 || duration.IsElapsed)
                    Error("PROJECT_FIXED_WORK_SCHEDULING", "Fixed-work tasks require stored work, positive assignment units, and a working duration.", task);
                else duration = ProjectDuration.WorkingMinutes(task.Work.Value.Minutes / units);
            }
            _nodes.Add(task, new Node { Task = task, Calendar = math, Elapsed = duration.IsElapsed,
                Minutes = checked(duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, _document)) });
        }
        foreach (var link in _document.Dependencies) {
            _token.ThrowIfCancellationRequested();
            if (link.CrossProject == true || link.Predecessor == null) { Error("PROJECT_EXTERNAL_SCHEDULING", "Resolve external dependencies explicitly before calculation.", link.Successor); continue; }
            if (!_nodes.TryGetValue(link.Predecessor, out var predecessor) || !_nodes.TryGetValue(link.Successor, out var successor)) {
                Error("PROJECT_DEPENDENCY_PROFILE", "Dependencies involving summaries, inactive tasks, or placeholders are outside this scheduling profile.", link.Successor); continue;
            }
            predecessor.Out.Add(link); successor.In.Add(link);
        }
        var degrees = _nodes.Values.ToDictionary(n => n, n => n.In.Count);
        var ready = new Queue<Node>(_nodes.Values.Where(n => degrees[n] == 0));
        while (ready.Count != 0) {
            _token.ThrowIfCancellationRequested(); var node = ready.Dequeue(); _order.Add(node);
            foreach (var link in node.Out) { var successor = _nodes[link.Successor]; if (--degrees[successor] == 0) ready.Enqueue(successor); }
        }
        if (_order.Count != _nodes.Count) Error("PROJECT_DEPENDENCY_CYCLE", "The schedule contains a dependency cycle.");
    }
    private DateTime Add(Node node, DateTime date, decimal minutes) => node.Elapsed
        ? date.AddTicks(checked((long)decimal.Round(minutes * TimeSpan.TicksPerMinute, 0, MidpointRounding.AwayFromZero))) : node.Calendar.Add(date, minutes);
    private DateTime Snap(Node node, DateTime date, bool forward) => node.Elapsed || node.Minutes == 0 ? date : node.Calendar.Snap(date, forward);
    private DateTime Lag(ProjectDependency link, DateTime date, bool reverse) {
        var predecessor = _nodes[link.Predecessor!]; var successor = _nodes[link.Successor];
        decimal minutes = link.LagPercent.HasValue ? predecessor.Minutes * link.LagPercent.Value / 100m : link.Lag.HasValue
            ? link.Lag.Value.Value * ProjectXmlValue.MinutesPerUnit(link.Lag.Value.Unit, link.Lag.Value.IsElapsed, _document) : 0m;
        if (reverse) minutes = -minutes;
        bool elapsed = link.Lag?.IsElapsed == true || (link.LagPercent.HasValue && predecessor.Elapsed);
        return elapsed ? date.AddTicks(checked((long)decimal.Round(minutes * TimeSpan.TicksPerMinute, 0, MidpointRounding.AwayFromZero))) : successor.Calendar.Add(date, minutes);
    }
    private static bool FromFinish(ProjectDependency link) => (link.Type ?? ProjectDependencyType.FinishToStart) is ProjectDependencyType.FinishToStart or ProjectDependencyType.FinishToFinish;
    private static bool ToFinish(ProjectDependency link) => (link.Type ?? ProjectDependencyType.FinishToStart) is ProjectDependencyType.FinishToFinish or ProjectDependencyType.StartToFinish;
    private static DateTime Max(DateTime first, DateTime second) => first > second ? first : second;
    private static DateTime Min(DateTime first, DateTime second) => first < second ? first : second;
    private ProjectCalendarMath CalendarMath(IEnumerable<ProjectCalendar> calendars) {
        var selected = calendars.Distinct().OrderBy(c => c.Uid).ToArray();
        string key = string.Join(":", selected.Select(c => c.Uid.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        if (!_calendars.TryGetValue(key, out var math)) {
            math = new ProjectCalendarMath(selected, _options.MaxCalendarDays, _token); _calendars.Add(key, math);
        }
        return math;
    }
}
