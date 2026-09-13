namespace OfficeIMO.Project;

/// <summary>A named, inclusive date range with weekday overrides. Unspecified weekdays inherit the ordinary calendar.</summary>
public sealed class ProjectWorkWeek : ProjectObject {
    internal ProjectWorkWeek(ProjectDocument document) : base(document) {
        WeekDays = new ProjectCollection<ProjectWeekDay>(document, () => new ProjectWeekDay(document), true, owner: this);
    }
    private string? _name;
    private DateTime? _from, _to;
    /// <summary>Optional descriptive label.</summary>
    public string? Name { get => _name; set => Set(ref _name, value); }
    /// <summary>Inclusive first date; null means no lower boundary.</summary>
    public DateTime? FromDate { get => _from; set => Set(ref _from, value, true); }
    /// <summary>Inclusive last date; null means no upper boundary.</summary>
    public DateTime? ToDate { get => _to; set => Set(ref _to, value, true); }
    /// <summary>Weekday overrides within this date range.</summary>
    public ProjectCollection<ProjectWeekDay> WeekDays { get; }
    /// <summary>Replaces one weekday's intervals. No intervals means non-working.</summary>
    public ProjectWeekDay SetWorkingDay(DayOfWeek day, params ProjectWorkingTime[] times) {
        if (!Enum.IsDefined(typeof(DayOfWeek), day)) throw new ArgumentOutOfRangeException(nameof(day));
        if (times == null) throw new ArgumentNullException(nameof(times));
        EnsureAttached();
        var item = WeekDays.FirstOrDefault(d => d.Day == day) ?? WeekDays.Add();
        item.Day = day; item.IsWorking = times.Length != 0;
        foreach (var old in item.WorkingTimes.ToArray()) item.WorkingTimes.Remove(old);
        foreach (var time in times) { var interval = item.WorkingTimes.Add(); interval.From = time.From; interval.To = time.To; }
        return item;
    }
}
