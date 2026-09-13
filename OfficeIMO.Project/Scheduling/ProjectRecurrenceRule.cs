using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Calendar-date pattern for a finite materialized task series, independent of working calendars.</summary>
public enum ProjectRecurrenceFrequency {
    /// <summary>Every selected number of calendar days.</summary>
    Daily,
    /// <summary>Selected weekdays in every selected number of weeks.</summary>
    Weekly,
    /// <summary>A day number or ordinal weekday in selected months.</summary>
    Monthly,
    /// <summary>A day number or ordinal weekday in the selected month of selected years.</summary>
    Yearly
}

/// <summary>
/// Finite Gregorian recurrence input. Expansion retains local wall time, skips nonexistent month days,
/// and never reads a native Project recurrence rule. Working calendars are applied later by scheduling.
/// </summary>
public sealed class ProjectRecurrenceRule {
    /// <summary>First eligible local date and time; its period anchors the interval.</summary>
    public DateTime Start { get; set; }
    /// <summary>Daily, weekly, monthly or yearly pattern.</summary>
    public ProjectRecurrenceFrequency Frequency { get; set; }
    /// <summary>Positive period stride.</summary>
    public int Interval { get; set; } = 1;
    /// <summary>Required finite number of occurrences, from 1 through 100000.</summary>
    public int Count { get; set; } = 1;
    /// <summary>Weekly selections. Empty selects Start's weekday. Other frequencies require an empty list.</summary>
    public IList<DayOfWeek> DaysOfWeek { get; } = new List<DayOfWeek>();
    /// <summary>Weekly period boundary, Monday by default.</summary>
    public DayOfWeek FirstDayOfWeek { get; set; } = DayOfWeek.Monday;
    /// <summary>Monthly/yearly day number. Null uses Start's day when no ordinal weekday is selected.</summary>
    public int? DayOfMonth { get; set; }
    /// <summary>Yearly month number. Null uses Start's month.</summary>
    public int? Month { get; set; }
    /// <summary>Monthly/yearly ordinal weekday, used together with WeekOrdinal.</summary>
    public DayOfWeek? WeekDay { get; set; }
    /// <summary>1 through 5 or -1 for the last matching weekday. Missing fifth weekdays are skipped.</summary>
    public int? WeekOrdinal { get; set; }

    /// <summary>Expands completely or fails without a partial result. MaxCalendarDays bounds the search, including dates skipped by the pattern.</summary>
    public IReadOnlyList<DateTime> Expand(int maxCalendarDays = 36600, CancellationToken cancellationToken = default) {
        ProjectCalendarMath.Local(Start);
        if (!Enum.IsDefined(typeof(ProjectRecurrenceFrequency), Frequency) || !Enum.IsDefined(typeof(DayOfWeek), FirstDayOfWeek) ||
            Interval < 1 || Interval > 366000 || Count < 1 || Count > 100000 || maxCalendarDays < 1 || maxCalendarDays > 366000)
            throw new ArgumentOutOfRangeException(nameof(ProjectRecurrenceRule));
        if (DaysOfWeek.Count > 7 || DaysOfWeek.Any(d => !Enum.IsDefined(typeof(DayOfWeek), d)) || DaysOfWeek.Distinct().Count() != DaysOfWeek.Count)
            throw new ArgumentException("Weekly recurrence requires distinct valid weekdays.");
        if (Frequency != ProjectRecurrenceFrequency.Weekly && DaysOfWeek.Count != 0 ||
            Frequency != ProjectRecurrenceFrequency.Yearly && Month.HasValue || Month < 1 || Month > 12 || DayOfMonth < 1 || DayOfMonth > 31 ||
            WeekDay.HasValue != WeekOrdinal.HasValue || WeekDay.HasValue && !Enum.IsDefined(typeof(DayOfWeek), WeekDay.Value) ||
            WeekOrdinal.HasValue && WeekOrdinal != -1 && (WeekOrdinal < 1 || WeekOrdinal > 5) ||
            WeekDay.HasValue && DayOfMonth.HasValue ||
            (Frequency == ProjectRecurrenceFrequency.Daily || Frequency == ProjectRecurrenceFrequency.Weekly) && (DayOfMonth.HasValue || WeekDay.HasValue))
            throw new ArgumentException("The recurrence pattern contains incompatible or invalid date selectors.");
        DateTime start = Start; var frequency = Frequency; int interval = Interval, count = Count;
        int month = Month ?? start.Month, day = DayOfMonth ?? start.Day;
        var weekday = WeekDay; var ordinal = WeekOrdinal;
        var weekly = new HashSet<DayOfWeek>(DaysOfWeek.Count == 0 ? new[] { start.DayOfWeek } : DaysOfWeek);
        int weekOffset = ((int)start.DayOfWeek - (int)FirstDayOfWeek + 7) % 7;
        var result = new List<DateTime>();
        for (int offset = 0; offset < maxCalendarDays; offset++) {
            cancellationToken.ThrowIfCancellationRequested();
            DateTime candidate = start.AddDays(offset);
            bool selected;
            if (frequency == ProjectRecurrenceFrequency.Daily) selected = offset % interval == 0;
            else if (frequency == ProjectRecurrenceFrequency.Weekly) selected = (offset + weekOffset) / 7 % interval == 0 && weekly.Contains(candidate.DayOfWeek);
            else {
                int period = frequency == ProjectRecurrenceFrequency.Monthly ? (candidate.Year - start.Year) * 12 + candidate.Month - start.Month : candidate.Year - start.Year;
                selected = period % interval == 0 && (frequency != ProjectRecurrenceFrequency.Yearly || candidate.Month == month);
                if (selected) selected = weekday.HasValue
                    ? candidate.DayOfWeek == weekday.Value && (ordinal == -1 ? candidate.Day + 7 > DateTime.DaysInMonth(candidate.Year, candidate.Month) : (candidate.Day - 1) / 7 + 1 == ordinal)
                    : candidate.Day == day;
            }
            if (!selected) continue;
            result.Add(candidate);
            if (result.Count == count) return new ReadOnlyCollection<DateTime>(result);
        }
        throw new InvalidOperationException("The recurrence search exceeded MaxCalendarDays before reaching Count.");
    }
}
