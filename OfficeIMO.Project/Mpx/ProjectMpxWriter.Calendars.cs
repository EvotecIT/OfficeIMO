namespace OfficeIMO.Project;

internal sealed partial class ProjectMpxWriter {
    private readonly Dictionary<ProjectCalendar, ProjectResource> _resourceCalendarOwners = new Dictionary<ProjectCalendar, ProjectResource>();
    private int _nextCalendarUid;
    private void Calendars() {
        var parents = new HashSet<ProjectCalendar>(_document.Calendars.Where(c => c.BaseCalendar != null).Select(c => c.BaseCalendar!));
        foreach (var group in _document.Resources.Where(r => r.Calendar != null).GroupBy(r => r.Calendar!)) {
            var calendar = group.Key;
            if (calendar.BaseCalendar != null && calendar != _document.Calendar && group.Count() == 1 && !parents.Contains(calendar))
                _resourceCalendarOwners.Add(calendar, group.Single());
        }
        if (_document.Calendars.Count - _resourceCalendarOwners.Count > 250)
            throw new NotSupportedException("MPX supports at most 250 emitted base calendars.");
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var calendar in _document.Calendars.Where(c => !_resourceCalendarOwners.ContainsKey(c))) {
            string name = calendar.Name ?? "Calendar " + calendar.Uid;
            if (calendar.Name == null) Diagnostic("PROJECT_MPX_CALENDAR_NAME_DEFAULT",
                "MPX requires a calendar name and normalizes an absent name to '" + name + "'.", Path(calendar, "Calendar") + "/Name");
            if (name.Length == 0 || !names.Add(name)) throw new InvalidDataException("MPX base calendars need nonempty, unique names.");
            _calendarNames.Add(calendar, name);
            WriteCalendar(calendar, name, false);
        }
    }
    private void ResourceCalendar(ProjectResource resource, string path) {
        if (resource.Calendar == null) return;
        var calendar = resource.Calendar; Handle(path + "/Calendar");
        if (_resourceCalendarOwners.TryGetValue(calendar, out var owner) && owner == resource) {
            string representedName = resource.Name ?? "Resource";
            if (!string.Equals(calendar.Name, representedName, StringComparison.Ordinal)) Diagnostic("PROJECT_MPX_RESOURCE_CALENDAR_NAME",
                "MPX derives the resource calendar name from the resource and uses 'Resource' when the resource name is absent.", Path(calendar, "Calendar") + "/Name");
            WriteCalendar(calendar, _calendarNames[calendar.BaseCalendar!], true);
        } else {
            Record("55", _calendarNames[calendar], "2", "2", "2", "2", "2", "2", "2");
            _nextCalendarUid++;
            Diagnostic("PROJECT_MPX_RESOURCE_CALENDAR_ID", "MPX represents a direct base-calendar binding by a resource calendar inheriting that base; a new calendar identity is assigned.", path + "/Calendar");
        }
    }
    private void WriteCalendar(ProjectCalendar calendar, string name, bool resource) {
            _token.ThrowIfCancellationRequested(); string path = Path(calendar, "Calendar");
            Handle(path + "/Name"); Handle(path + "/Uid");
            if (calendar.Uid != ++_nextCalendarUid) Diagnostic("PROJECT_MPX_CALENDAR_ID", "MPX identifies calendars by name; numeric calendar identities are assigned again when reopened.", path + "/Uid");
            Handle(path + "/IsBaseCalendar"); Handle(path + "/BaseCalendar");
            if (!resource && (calendar.BaseCalendar != null || calendar.IsBaseCalendar == false))
                Diagnostic("PROJECT_MPX_CALENDAR_FLATTENED", "Calendar inheritance is flattened into a named base calendar with effective working times.", path);
            var math = new ProjectCalendarMath(new[] { calendar }, 366000, _token);
            var header = new List<string?> { resource ? "55" : "20", name };
            var patterns = new List<string[]>();
            for (int day = 0; day < 7; day++) {
                var pattern = Intervals(math.OrdinaryDayPattern((DayOfWeek)day)); patterns.Add(pattern);
                bool inherit = resource && !calendar.WeekDays.Any(d => d.Day == (DayOfWeek)day && d.IsWorking.HasValue);
                header.Add(inherit ? "2" : pattern.Length == 0 ? "0" : "1");
            }
            Record(header.ToArray());
            for (int day = 0; day < 7; day++) if (header[day + 2] != "2") Record(new[] { resource ? "56" : "25", ProjectMpxValues.Text(day + 1) }.Concat(patterns[day]).ToArray());
            var dates = new SortedSet<DateTime>();
            void Range(DateTime? from, DateTime? to) {
                if (!from.HasValue || !to.HasValue || to.Value.Date < from.Value.Date) throw new NotSupportedException("MPX calendar overrides need an ordered, finite date range.");
                int count = checked((int)(to.Value.Date - from.Value.Date).TotalDays + 1);
                if (count > 366000) throw new NotSupportedException("MPX calendar expansion exceeds 366000 days.");
                for (int i = 0; i < count; i++) { _token.ThrowIfCancellationRequested(); dates.Add(from.Value.Date.AddDays(i)); if (dates.Count > 366000) throw new NotSupportedException("MPX calendar expansion exceeds 366000 days."); }
            }
            for (var current = calendar; current != null; current = current.BaseCalendar) {
                if (!resource || current == calendar) {
                    foreach (var exception in current.Exceptions) Range(exception.FromDate, exception.ToDate);
                    foreach (var day in current.WeekDays.Where(d => d.Day == null)) Range(day.FromDate, day.ToDate);
                }
                foreach (var week in current.WorkWeeks) Range(week.FromDate, week.ToDate);
            }
            DateTime? rangeFrom = null, rangeTo = null; string[]? rangePattern = null; int exceptions = 0;
            void Flush() {
                if (!rangeFrom.HasValue) return;
                if (++exceptions > 250) throw new NotSupportedException("MPX calendar projection exceeds 250 exception records.");
                Record(new[] { resource ? "57" : "26", rangeFrom.Value.ToString("yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture), rangeTo!.Value.ToString("yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture), rangePattern!.Length == 0 ? "0" : "1" }.Concat(rangePattern).ToArray());
            }
            foreach (var date in dates) {
                _token.ThrowIfCancellationRequested(); var pattern = Intervals(math.OwnerDayPattern(date));
                if (rangeTo.HasValue && date.Ticks - rangeTo.Value.Ticks == TimeSpan.TicksPerDay && pattern.SequenceEqual(rangePattern!)) rangeTo = date;
                else { Flush(); rangeFrom = rangeTo = date; rangePattern = pattern; }
            }
            Flush();
            HandleTree(path + "/Day"); HandleTree(path + "/Exception"); HandleTree(path + "/Week");
            if (calendar.WorkWeeks.Count != 0) Diagnostic("PROJECT_MPX_WORK_WEEK_FLATTENED", "Bounded work weeks become date exceptions; labels and work-week structure are omitted.", path + "/Week");
            if (calendar.Exceptions.Any(e => e.Name != null)) Diagnostic("PROJECT_MPX_EXCEPTION_LABEL_LOSS", "Calendar exception dates and working times are retained; MPX has no exception labels.", path + "/Exception");
    }
    private static string[] Intervals(IReadOnlyList<ProjectWorkingInterval> intervals) {
        if (intervals.Count > 3) throw new NotSupportedException("MPX allows at most three working intervals per day.");
        var result = new List<string>();
        foreach (var interval in intervals) {
            if (!interval.From.HasValue || !interval.To.HasValue) throw new InvalidDataException("MPX working intervals need two endpoints.");
            if (interval.From.Value.Ticks % TimeSpan.TicksPerMinute != 0 || interval.To.Value.Ticks % TimeSpan.TicksPerMinute != 0)
                throw new NotSupportedException("MPX working intervals require whole minutes.");
            result.Add(ProjectMpxValues.Text(interval.From)); result.Add(ProjectMpxValues.Text(interval.To));
        }
        return result.ToArray();
    }
}
