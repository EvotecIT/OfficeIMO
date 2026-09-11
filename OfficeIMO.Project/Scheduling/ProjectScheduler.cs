namespace OfficeIMO.Project;

internal sealed partial class ProjectScheduler {
    private readonly ProjectDocument _document;
    private readonly ProjectScheduleOptions _options;
    private readonly CancellationToken _token;
    private readonly long _revision;
    private readonly IReadOnlyDictionary<int, DateTime>? _notBefore;
    private readonly IReadOnlyDictionary<int, IReadOnlyList<ProjectWorkingRange>>? _splits;
    private readonly ProjectExternalProjectContext _externalContext;
    private readonly Dictionary<ProjectDependency, (ProjectTaskSchedule Task, ProjectDocument Document)> _externalDependencies = new();
    private readonly Dictionary<ProjectDocument, ProjectExternalScheduleSource> _externalSources = new();
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
        internal DateTime EarlyAnchor, LateAnchor, PlanAnchor, BeforeLevelingAnchor;
        internal ProjectTaskAllocation? Allocation;
        internal ProjectTaskAllocation.Result? Plan;
        internal readonly List<ProjectDependency> In = new List<ProjectDependency>(), Out = new List<ProjectDependency>();
    }
    internal ProjectScheduler(ProjectDocument document, ProjectScheduleOptions options, CancellationToken token, IReadOnlyDictionary<int, DateTime>? notBefore = null, ProjectExternalProjectContext? externalContext = null,
        IReadOnlyDictionary<int, IReadOnlyList<ProjectWorkingRange>>? splits = null) {
        _notBefore = notBefore;
        _splits = splits;
        _document = document; _options = new ProjectScheduleOptions { MaxCalendarDays = options.MaxCalendarDays, CriticalSlackMinutes = options.CriticalSlackMinutes,
            CalculateAssignments = options.CalculateAssignments, RedistributeEffortDrivenWork = options.RedistributeEffortDrivenWork,
            RescheduleRemainingAfterStatusDate = options.RescheduleRemainingAfterStatusDate, RecalculateActualCosts = options.RecalculateActualCosts, MaxIntervals = options.MaxIntervals,
            ExternalProjectResolver = options.ExternalProjectResolver, MaxExternalProjects = options.MaxExternalProjects, MaxExternalDepth = options.MaxExternalDepth, MaxExternalTasks = options.MaxExternalTasks };
        _token = token; _revision = document.Revision;
        _externalContext = externalContext ?? new ProjectExternalProjectContext(_options, token);
    }
    private void Error(string code, string message, ProjectTask? task = null) => _diagnostics.Add(new ProjectDiagnostic(code,
        ProjectDiagnosticSeverity.Error, message, task == null ? "/Project" : "/Task[UID=" + task.Uid + "]"));
    private bool Failed => _diagnostics.Any(d => d.Severity == ProjectDiagnosticSeverity.Error);
    internal ProjectScheduleResult Calculate() {
        _externalContext.Enter(_document);
        try { return CalculateCore(); }
        finally { _externalContext.Leave(_document); }
    }
    private ProjectScheduleResult CalculateCore() {
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
                node.PlanAnchor = late ? node.LateAnchor : node.EarlyAnchor;
            }
            CheckFinalBounds();
            return Result(BuildResults(horizon));
        } catch (Exception exception) when (exception is InvalidDataException || exception is IOException || exception is InvalidOperationException ||
            exception is NotSupportedException || exception is ArgumentException || exception is OverflowException) {
            Error("PROJECT_CALCULATION_UNSUPPORTED", exception.Message); return Result();
        }
    }
    private ProjectScheduleResult Result(IEnumerable<ProjectTaskSchedule>? tasks = null) {
        _token.ThrowIfCancellationRequested();
        if (_document.Revision != _revision) throw new InvalidOperationException("The document changed during schedule calculation.");
        var calculated = tasks?.ToArray() ?? Array.Empty<ProjectTaskSchedule>();
        if (!_options.CalculateAssignments) CheckStoredAssignmentDates(calculated);
        var assignments = _order.Where(n => n.Plan != null).SelectMany(n => n.Plan!.Assignments).ToArray();
        if (!_externalContext.WithinIntervalLimit(_document, assignments.Sum(a => (long)a.Intervals.Count + a.Costs.Count)))
            Error("PROJECT_CALCULATION_INTERVAL_LIMIT", "The calculated local and external assignment and cost intervals exceed MaxIntervals.");
        foreach (var source in _externalSources.Values) source.ValidateCurrent();
        return new ProjectScheduleResult(_document, _revision, calculated, _diagnostics, assignments, _options.CalculateAssignments, _externalSources.Values);
    }
    private void Prepare() {
        var all = _document.AllTasks.ToArray();
        CheckSourceProfile();
        var allAssignments = _document.Assignments.Where(a => a.Task != null).GroupBy(a => a.Task!).ToDictionary(g => g.Key, g => g.ToArray());
        var assignments = _document.Assignments.Where(a => a.Task != null && a.Resource?.Type == ProjectResourceType.Work)
            .GroupBy(a => a.Task!).ToDictionary(g => g.Key, g => g.ToArray());
        foreach (var task in all) {
            _token.ThrowIfCancellationRequested();
            if (task.IsSummary || task.IsNull == true || task.IsActive == false || task.Uid == 0) continue;
            var calendar = task.Calendar ?? _document.Calendar;
            if (calendar == null) { Error("PROJECT_CALCULATION_CALENDAR", "Set an explicit task or project calendar.", task); continue; }
            if (task.IsManual != true && !task.Duration.HasValue) { Error("PROJECT_CALCULATION_DURATION", "Automatic tasks require an explicit duration.", task); continue; }
            if (task.IsManual == true && (!task.Start.HasValue || !task.Finish.HasValue)) { Error("PROJECT_MANUAL_DATES", "Manual tasks require both stored dates.", task); continue; }
            if (!_options.CalculateAssignments && (task.ActualStart.HasValue || task.ActualFinish.HasValue || task.ActualWork?.Minutes > 0 || task.ActualDuration?.Value > 0 || task.PercentComplete > 0 || task.PercentWorkComplete > 0))
                Error("PROJECT_PROGRESS_SCHEDULING", "Progress and status-date rescheduling are not qualified by the date-only scheduling profile.", task);
            var calendars = new List<ProjectCalendar> { calendar };
            assignments.TryGetValue(task, out var assigned);
            if (assigned != null && !_options.CalculateAssignments) {
                var resourceCalendars = assigned.Select(a => a.Resource!.Calendar ?? _document.Calendar ?? calendar).Distinct().ToArray();
                if (resourceCalendars.Length > 1) Error("PROJECT_RESOURCE_CALENDAR_SCHEDULING", "Assignments with different resource calendars require independent assignment scheduling.", task);
                if (task.Calendar == null) calendars.Clear();
                calendars.AddRange(resourceCalendars);
            }
            var math = CalendarMath(calendars);
            var duration = task.Duration ?? ProjectDuration.WorkingMinutes(math.Between(task.Start!.Value, task.Finish!.Value));
            if (task.Type == ProjectTaskType.FixedWork && task.IsManual != true && !_options.CalculateAssignments) {
                decimal units = assigned?.Sum(a => a.Units?.Value ?? 0m) ?? 0m;
                if (!task.Work.HasValue || units <= 0 || duration.IsElapsed)
                    Error("PROJECT_FIXED_WORK_SCHEDULING", "Fixed-work tasks require stored work, positive assignment units, and a working duration.", task);
                else duration = ProjectDuration.WorkingMinutes(task.Work.Value.Minutes / units);
            }
            var node = new Node { Task = task, Calendar = math, Elapsed = duration.IsElapsed,
                Minutes = checked(duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, _document)) };
            if (_options.CalculateAssignments) {
                node.Allocation = new ProjectTaskAllocation(task, allAssignments.TryGetValue(task, out var current) ? current : Array.Empty<ProjectAssignment>(),
                    _options, CalendarMath, d => _diagnostics.Add(d), _token,
                    _splits != null && _splits.TryGetValue(task.Uid, out var taskSplits) ? taskSplits : null);
                if (node.Allocation.HasActuals && _document.Settings.ScheduleFromStart == false)
                    Error("PROJECT_PROGRESS_BACKWARD", "Recorded progress requires forward remaining-work scheduling.", task);
            }
            _nodes.Add(task, node);
        }
        foreach (var link in _document.Dependencies) {
            _token.ThrowIfCancellationRequested();
            if (link.CrossProject == true) {
                if (!_nodes.TryGetValue(link.Successor, out var local)) { Error("PROJECT_DEPENDENCY_PROFILE", "External dependencies require an active local successor.", link.Successor); continue; }
                var resolved = _externalContext.Resolve(_document, link);
                _externalDependencies.Add(link, (resolved.Task, resolved.Schedule.Document));
                if (!_externalSources.ContainsKey(resolved.Schedule.Document)) foreach (var diagnostic in resolved.Schedule.Report.Diagnostics.Where(d => d.Severity != ProjectDiagnosticSeverity.Error))
                    _diagnostics.Add(new ProjectDiagnostic(diagnostic.Code, diagnostic.Severity, "External project: " + diagnostic.Message,
                        "/ExternalProject[" + resolved.Reference + "]" + diagnostic.Location, diagnostic.RepresentsLoss));
                _externalSources[resolved.Schedule.Document] = new ProjectExternalScheduleSource(resolved.Reference, resolved.Schedule.Document, resolved.Schedule.ModelRevision);
                foreach (var source in resolved.Schedule.ExternalSources) _externalSources[source.Document] = source;
                local.In.Add(link); continue;
            }
            if (link.Predecessor == null) { Error("PROJECT_EXTERNAL_SCHEDULING", "The dependency has no resolved predecessor.", link.Successor); continue; }
            if (!_nodes.TryGetValue(link.Predecessor, out var predecessor) || !_nodes.TryGetValue(link.Successor, out var successor)) {
                Error("PROJECT_DEPENDENCY_PROFILE", "Dependencies involving summaries, inactive tasks, or placeholders are outside this scheduling profile.", link.Successor); continue;
            }
            predecessor.Out.Add(link); successor.In.Add(link);
        }
        var degrees = _nodes.Values.ToDictionary(n => n, n => n.In.Count(l => !_externalDependencies.ContainsKey(l)));
        var ready = new Queue<Node>(_nodes.Values.Where(n => degrees[n] == 0));
        while (ready.Count != 0) {
            _token.ThrowIfCancellationRequested(); var node = ready.Dequeue(); _order.Add(node);
            foreach (var link in node.Out) { var successor = _nodes[link.Successor]; if (--degrees[successor] == 0) ready.Enqueue(successor); }
        }
        if (_order.Count != _nodes.Count) Error("PROJECT_DEPENDENCY_CYCLE", "The schedule contains a dependency cycle.");
        if (_externalSources.Count > 0) _diagnostics.Add(new ProjectDiagnostic("PROJECT_EXTERNAL_SCHEDULE_BOUNDARIES", ProjectDiagnosticSeverity.Information,
            "External predecessor dates were calculated through the caller resolver. Task, assignment, slack, and resource-capacity results remain scoped to the local project.", "/Project"));
    }
    private DateTime Add(Node node, DateTime date, decimal minutes) {
        if (node.Allocation != null && minutes != 0 && Math.Abs(minutes) == node.Minutes) {
            var result = node.Allocation.Build(date, minutes > 0); return minutes > 0 ? result.Finish : result.Start;
        }
        return node.Elapsed ? date.AddTicks(checked((long)decimal.Round(minutes * TimeSpan.TicksPerMinute, 0, MidpointRounding.AwayFromZero))) : node.Calendar.Add(date, minutes);
    }
    private DateTime Snap(Node node, DateTime date, bool forward) => node.Elapsed || node.Minutes == 0 ? date : node.Calendar.Snap(date, forward);
    private DateTime Lag(ProjectDependency link, DateTime date, bool reverse) {
        var successor = _nodes[link.Successor];
        decimal predecessorMinutes; bool predecessorElapsed;
        if (_externalDependencies.TryGetValue(link, out var external)) {
            var duration = external.Task.Duration; predecessorElapsed = duration.IsElapsed;
            predecessorMinutes = duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, external.Document);
        } else { var predecessor = _nodes[link.Predecessor!]; predecessorMinutes = predecessor.Minutes; predecessorElapsed = predecessor.Elapsed; }
        decimal minutes = link.LagPercent.HasValue ? predecessorMinutes * link.LagPercent.Value / 100m : link.Lag.HasValue
            ? link.Lag.Value.Value * ProjectXmlValue.MinutesPerUnit(link.Lag.Value.Unit, link.Lag.Value.IsElapsed, _document) : 0m;
        if (reverse) minutes = -minutes;
        bool elapsed = link.Lag?.IsElapsed == true || (link.LagPercent.HasValue && predecessorElapsed);
        return elapsed ? date.AddTicks(checked((long)decimal.Round(minutes * TimeSpan.TicksPerMinute, 0, MidpointRounding.AwayFromZero))) : successor.Calendar.Add(date, minutes);
    }
    private (DateTime Start, DateTime Finish) PredecessorBounds(ProjectDependency link, bool early) {
        if (_externalDependencies.TryGetValue(link, out var external)) return (external.Task.Start, external.Task.Finish);
        var node = _nodes[link.Predecessor!]; return early ? (node.EarlyStart, node.EarlyFinish) : (node.Start, node.Finish);
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
