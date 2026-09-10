namespace OfficeIMO.Project;

/// <summary>A base or derived calendar with stored working days and exceptions; no scheduling runs during editing.</summary>
public sealed class ProjectCalendar : ProjectNamedEntity {
    internal ProjectCalendar(ProjectDocument document, int uid) : base(document, uid) {
        WeekDays = new ProjectCollection<ProjectWeekDay>(document, () => new ProjectWeekDay(document), true, owner: this);
        Exceptions = new ProjectCollection<ProjectCalendarException>(document, () => new ProjectCalendarException(document), true, owner: this);
    }
    private ProjectCalendar? _baseCalendar;
    /// <summary>The inherited calendar, or null for a base calendar.</summary>
    public ProjectCalendar? BaseCalendar {
        get => _baseCalendar;
        set {
            CheckReference(value);
            var visited = new HashSet<ProjectCalendar>();
            for (var current = value; current != null; current = current.BaseCalendar)
                if (current == this || !visited.Add(current)) throw new ArgumentException("Calendar inheritance must not contain a cycle.");
            Set(ref _baseCalendar, value, true);
            if (!Document.Loading) SourceBaseCalendarUid = null;
        }
    }
    private bool? _isBaseCalendar;
    /// <summary>Source base-calendar flag; null preserves absence.</summary>
    public bool? IsBaseCalendar { get => _isBaseCalendar; set => Set(ref _isBaseCalendar, value, true); }
    internal int? SourceBaseCalendarUid { get; set; }
    /// <summary>Explicit day-of-week declarations; absent days may be inherited.</summary>
    public ProjectCollection<ProjectWeekDay> WeekDays { get; }
    /// <summary>Date exceptions, including retained unmodeled recurrence data.</summary>
    public ProjectCollection<ProjectCalendarException> Exceptions { get; }

    /// <summary>Replaces the working intervals for a day without changing other calendar days.</summary>
    public ProjectWeekDay SetWorkingDay(DayOfWeek day, params ProjectWorkingTime[] times) {
        if (!Enum.IsDefined(typeof(DayOfWeek), day)) throw new ArgumentOutOfRangeException(nameof(day));
        if (times == null) throw new ArgumentNullException(nameof(times));
        EnsureAttached();
        var item = WeekDays.FirstOrDefault(d => d.Day == day) ?? WeekDays.Add();
        item.Day = day;
        item.IsWorking = times.Length != 0;
        foreach (var old in item.WorkingTimes.ToArray()) item.WorkingTimes.Remove(old);
        foreach (var time in times) {
            var range = item.WorkingTimes.Add(); range.From = time.From; range.To = time.To;
        }
        return item;
    }
}

/// <summary>A working interval in a day's local clock, without host-timezone conversion.</summary>
public readonly struct ProjectWorkingTime {
    /// <summary>Creates an interval. An end earlier than the start denotes an overnight interval.</summary>
    public ProjectWorkingTime(TimeSpan from, TimeSpan to) {
        if (from < TimeSpan.Zero || from >= TimeSpan.FromDays(1)) throw new ArgumentOutOfRangeException(nameof(from));
        if (to < TimeSpan.Zero || to >= TimeSpan.FromDays(1)) throw new ArgumentOutOfRangeException(nameof(to));
        if (from == to) throw new ArgumentException("A working interval must have nonzero duration.");
        From = from; To = to;
    }
    /// <summary>Start on the local clock.</summary>
    public TimeSpan From { get; }
    /// <summary>End on the local clock.</summary>
    public TimeSpan To { get; }
    /// <summary>Creates an interval from whole hours.</summary>
    public static ProjectWorkingTime Hours(int from, int to) => new ProjectWorkingTime(TimeSpan.FromHours(from), TimeSpan.FromHours(to));
}

/// <summary>A mutable working interval retained in its enclosing day or exception.</summary>
public sealed class ProjectWorkingInterval : ProjectObject {
    internal ProjectWorkingInterval(ProjectDocument document) : base(document) { }
    private TimeSpan? _from, _to;
    /// <summary>Start on the local clock, or null for an absent source value.</summary>
    public TimeSpan? From { get => _from; set => Set(ref _from, value, true); }
    /// <summary>End on the local clock, or null for an absent source value.</summary>
    public TimeSpan? To { get => _to; set => Set(ref _to, value, true); }
}

/// <summary>An explicit weekday declaration; a missing weekday is different from a non-working one.</summary>
public sealed class ProjectWeekDay : ProjectObject {
    internal ProjectWeekDay(ProjectDocument document) : base(document) {
        WorkingTimes = new ProjectCollection<ProjectWorkingInterval>(document, () => new ProjectWorkingInterval(document), true, owner: this);
    }
    private DayOfWeek? _day;
    private bool? _isWorking;
    private DateTime? _fromDate, _toDate;
    /// <summary>Weekday, or null for a preserved special/source day definition.</summary>
    public DayOfWeek? Day { get => _day; set => Set(ref _day, value, true); }
    /// <summary>Explicit working status; null preserves absence.</summary>
    public bool? IsWorking { get => _isWorking; set => Set(ref _isWorking, value, true); }
    /// <summary>Start of a legacy date exception when Day is null; absent for ordinary weekdays.</summary>
    public DateTime? FromDate { get => _fromDate; set => Set(ref _fromDate, value, true); }
    /// <summary>End of a legacy date exception when Day is null; absent for ordinary weekdays.</summary>
    public DateTime? ToDate { get => _toDate; set => Set(ref _toDate, value, true); }
    /// <summary>Working clock intervals.</summary>
    public ProjectCollection<ProjectWorkingInterval> WorkingTimes { get; }
}

/// <summary>A date-range calendar exception. Additional recurrence metadata is preserved by XML save.</summary>
public sealed class ProjectCalendarException : ProjectObject {
    internal ProjectCalendarException(ProjectDocument document) : base(document) {
        WorkingTimes = new ProjectCollection<ProjectWorkingInterval>(document, () => new ProjectWorkingInterval(document), true, owner: this);
    }
    private string? _name;
    private DateTime? _fromDate, _toDate;
    private bool? _isWorking;
    /// <summary>Exception label.</summary>
    public string? Name { get => _name; set => Set(ref _name, value); }
    /// <summary>Inclusive first date.</summary>
    public DateTime? FromDate { get => _fromDate; set => Set(ref _fromDate, value, true); }
    /// <summary>Inclusive last date.</summary>
    public DateTime? ToDate { get => _toDate; set => Set(ref _toDate, value, true); }
    /// <summary>Whether the exception supplies working intervals.</summary>
    public bool? IsWorking { get => _isWorking; set => Set(ref _isWorking, value, true); }
    /// <summary>Explicit working intervals.</summary>
    public ProjectCollection<ProjectWorkingInterval> WorkingTimes { get; }
}
