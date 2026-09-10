namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private void ValidateCalendars(Finding add, CancellationToken token) {
        foreach (var calendar in Calendars) {
            token.ThrowIfCancellationRequested();
            string location = "/Calendar[UID=" + calendar.Uid + "]";
            if (calendar.BaseCalendar == null && calendar.SourceBaseCalendarUid > 0)
                add("PROJECT_CALENDAR_REFERENCE", "The calendar references a missing base calendar.", location);
            var days = new HashSet<DayOfWeek>();
            foreach (var day in calendar.WeekDays) {
                if (day.Day.HasValue && (!Enum.IsDefined(typeof(DayOfWeek), day.Day.Value) || !days.Add(day.Day.Value)))
                    add("PROJECT_CALENDAR_DAY", "Explicit weekdays must be valid and unique.", location);
                CheckIntervals(day.WorkingTimes, location, add);
                CheckDateRange(day.FromDate, day.ToDate, location + "/WeekDay/TimePeriod", add);
                if (day.Day == null && (!day.FromDate.HasValue || !day.ToDate.HasValue))
                    add("PROJECT_CALENDAR_PERIOD", "A legacy exception day needs a complete date range.", location);
            }
            foreach (var exception in calendar.Exceptions) {
                token.ThrowIfCancellationRequested();
                CheckDateRange(exception.FromDate, exception.ToDate, location + "/Exception", add);
                CheckIntervals(exception.WorkingTimes, location + "/Exception", add);
            }
        }
    }
    private static void CheckIntervals(ProjectCollection<ProjectWorkingInterval> intervals, string location, Finding add) {
        foreach (var interval in intervals) {
            if (!interval.From.HasValue || !interval.To.HasValue || interval.From < TimeSpan.Zero || interval.From >= TimeSpan.FromDays(1) ||
                interval.To < TimeSpan.Zero || interval.To >= TimeSpan.FromDays(1) || interval.From == interval.To)
                add("PROJECT_WORKING_INTERVAL", "Working intervals need distinct clock times within a day.", location);
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
