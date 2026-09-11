namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private void ValidateCalendars(Finding add, CancellationToken token) {
        foreach (var calendar in Calendars) {
            token.ThrowIfCancellationRequested();
            string location = "/Calendar[UID=" + calendar.Uid + "]";
            if (calendar.IsBaseCalendar == true && calendar.BaseCalendar != null)
                add("PROJECT_CALENDAR_KIND", "A base calendar cannot also inherit a resource base-calendar reference. Create a derived calendar with Calendars.Add(name, baseCalendar).", location);
            if (calendar.BaseCalendar == null && calendar.SourceBaseCalendarUid > 0)
                add("PROJECT_CALENDAR_REFERENCE", "The calendar references a missing base calendar.", location);
            var days = new HashSet<DayOfWeek>();
            foreach (var day in calendar.WeekDays) {
                token.ThrowIfCancellationRequested();
                if (day.Day.HasValue && (!Enum.IsDefined(typeof(DayOfWeek), day.Day.Value) || !days.Add(day.Day.Value)))
                    add("PROJECT_CALENDAR_DAY", "Explicit weekdays must be valid and unique.", location);
                CheckIntervals(day.WorkingTimes, location, add, token);
                CheckWorkingStatus(day.IsWorking, day.WorkingTimes.Count, location + "/WeekDay", add);
                CheckDateRange(day.FromDate, day.ToDate, location + "/WeekDay/TimePeriod", add);
                if (day.Day == null && (!day.FromDate.HasValue || !day.ToDate.HasValue))
                    add("PROJECT_CALENDAR_PERIOD", "A legacy exception day needs a complete date range.", location);
            }
            foreach (var exception in calendar.Exceptions) {
                token.ThrowIfCancellationRequested();
                CheckDateRange(exception.FromDate, exception.ToDate, location + "/Exception", add);
                CheckIntervals(exception.WorkingTimes, location + "/Exception", add, token);
                CheckWorkingStatus(exception.IsWorking, exception.WorkingTimes.Count, location + "/Exception", add);
            }
            foreach (var week in calendar.WorkWeeks) {
                token.ThrowIfCancellationRequested();
                CheckDateRange(week.FromDate, week.ToDate, location + "/WorkWeek", add);
                var weekdays = new HashSet<DayOfWeek>();
                foreach (var day in week.WeekDays) {
                    if (!day.Day.HasValue || !Enum.IsDefined(typeof(DayOfWeek), day.Day.Value) || !weekdays.Add(day.Day.Value))
                        add("PROJECT_CALENDAR_DAY", "Work-week overrides require unique valid weekdays.", location + "/WorkWeek");
                    CheckIntervals(day.WorkingTimes, location + "/WorkWeek", add, token);
                    CheckWorkingStatus(day.IsWorking, day.WorkingTimes.Count, location + "/WorkWeek", add);
                }
            }
            CheckCalendarOverlaps(calendar.WorkWeeks.Select(w => (w.FromDate, w.ToDate)), location + "/WorkWeek", add, token);
            CheckCalendarOverlaps(calendar.Exceptions.Select(e => (e.FromDate, e.ToDate))
                .Concat(calendar.WeekDays.Where(d => d.Day == null).Select(d => (d.FromDate, d.ToDate))), location + "/Exception", add, token);
        }
    }
    private static void CheckCalendarOverlaps(IEnumerable<(DateTime? From, DateTime? To)> periods, string location, Finding add, CancellationToken token) {
        DateTime? previousEnd = null;
        foreach (var period in periods.OrderBy(p => p.From?.Date ?? DateTime.MinValue)) {
            token.ThrowIfCancellationRequested();
            var start = period.From?.Date ?? DateTime.MinValue;
            var finish = period.To?.Date ?? DateTime.MaxValue.Date;
            if (previousEnd.HasValue && start <= previousEnd.Value)
                add("PROJECT_CALENDAR_OVERLAP", "Calendar overrides at the same precedence cannot overlap inclusive dates.", location);
            if (!previousEnd.HasValue || finish > previousEnd.Value) previousEnd = finish;
        }
    }
    private static void CheckWorkingStatus(bool? working, int intervalCount, string location, Finding add) {
        if (working != true && intervalCount != 0)
            add("PROJECT_CALENDAR_WORKING_STATUS", "Working intervals require an explicitly working day.", location);
        else if (working == true && intervalCount == 0)
            add("PROJECT_CALENDAR_WORKING_STATUS", "An explicitly working day requires at least one working interval.", location);
    }
    private static void CheckIntervals(ProjectCollection<ProjectWorkingInterval> intervals, string location, Finding add, CancellationToken token) {
        foreach (var interval in intervals) {
            token.ThrowIfCancellationRequested();
            if (!interval.From.HasValue || !interval.To.HasValue || interval.From < TimeSpan.Zero || interval.From >= TimeSpan.FromDays(1) ||
                interval.To < TimeSpan.Zero || interval.To >= TimeSpan.FromDays(1))
                add("PROJECT_WORKING_INTERVAL", "Working intervals need clock times within a day.", location);
        }
    }
    private void ValidateDependencies(Finding add, CancellationToken token) {
        var outgoing = new Dictionary<ProjectTask, List<ProjectTask>>();
        var indegree = TaskIndex.Values.ToDictionary(t => t, _ => 0);
        var edges = new HashSet<string>(StringComparer.Ordinal);
        foreach (var link in Dependencies) {
            token.ThrowIfCancellationRequested();
            string location = "/Task[UID=" + link.Successor.Uid + "]/PredecessorLink";
            CheckEnum(link.Type, location + "/Type", add);
            try {
                decimal? rawLag = ProjectXmlCodec.DependencyLag(link, this);
                if (rawLag.HasValue && rawLag.Value != decimal.Truncate(rawLag.Value))
                    add("PROJECT_LAG_PRECISION", "MSPDI stores lag as integer tenths of a minute or integer percent; saving rounds midpoint values away from zero.", location + "/LinkLag", ProjectDiagnosticSeverity.Warning, true);
            } catch (OverflowException) {
                add("PROJECT_LAG_RANGE", "The dependency lag exceeds the representable range.", location + "/LinkLag");
            }
            if (link.Predecessor == null) {
                add("PROJECT_PREDECESSOR_REFERENCE", link.CrossProject == true ? "An external predecessor is preserved but not resolved or calculated." : "The predecessor task is missing.", location,
                    link.CrossProject == true ? ProjectDiagnosticSeverity.Warning : ProjectDiagnosticSeverity.Error);
                continue;
            }
            if (link.Predecessor == link.Successor) add("PROJECT_DEPENDENCY_SELF", "A task cannot depend on itself.", location);
            if (!edges.Add(link.Predecessor.Uid + ":" + link.Successor.Uid)) add("PROJECT_DEPENDENCY_DUPLICATE", "A predecessor/successor relationship is duplicated.", location);
            if (!outgoing.TryGetValue(link.Predecessor, out var successors)) { successors = new List<ProjectTask>(); outgoing.Add(link.Predecessor, successors); }
            successors.Add(link.Successor); indegree[link.Successor]++;
        }
        var queue = new Queue<ProjectTask>(indegree.Where(p => p.Value == 0).Select(p => p.Key));
        int visited = 0;
        while (queue.Count != 0) {
            token.ThrowIfCancellationRequested();
            var task = queue.Dequeue(); visited++;
            if (!outgoing.TryGetValue(task, out var successors)) continue;
            foreach (var next in successors) if (--indegree[next] == 0) queue.Enqueue(next);
        }
        if (visited != indegree.Count) add("PROJECT_DEPENDENCY_CYCLE", "Task dependencies contain a cycle.", "/Project/Tasks");
    }
}
