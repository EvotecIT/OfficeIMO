namespace OfficeIMO.Project;

/// <summary>A half-open working interval on the project's local, timezone-free timeline.</summary>
public readonly struct ProjectWorkingRange {
    internal ProjectWorkingRange(DateTime start, DateTime finish) { Start = start; Finish = finish; }
    /// <summary>Inclusive interval start.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive interval finish.</summary>
    public DateTime Finish { get; }
}

public sealed partial class ProjectCalendar {
    /// <summary>Returns effective intervals within one local date, including an overnight shift started the preceding day.</summary>
    public IReadOnlyList<ProjectWorkingRange> GetWorkingIntervals(DateTime date, CancellationToken cancellationToken = default) =>
        new ProjectCalendarMath(new[] { this }, 36600, cancellationToken).Day(date).AsReadOnly();

    /// <summary>Adds signed working minutes, skipping non-working intervals. Uses local wall time without host timezone or DST conversion.</summary>
    public DateTime AddWorkingMinutes(DateTime start, decimal minutes, int maxCalendarDays = 36600, CancellationToken cancellationToken = default) =>
        new ProjectCalendarMath(new[] { this }, maxCalendarDays, cancellationToken).Add(start, minutes);

    /// <summary>Counts signed working minutes between local dates. Split shifts are counted once, even when intervals overlap.</summary>
    public decimal WorkingMinutesBetween(DateTime start, DateTime finish, int maxCalendarDays = 36600, CancellationToken cancellationToken = default) =>
        new ProjectCalendarMath(new[] { this }, maxCalendarDays, cancellationToken).Between(start, finish);
}

/// <summary>One calculation's bounded calendar intersection, shared by duration, dependency, and resource analysis.</summary>
internal sealed class ProjectCalendarMath {
    private readonly ProjectCalendar[] _calendars;
    private readonly int _maxDays;
    private readonly CancellationToken _token;
    private readonly Dictionary<DateTime, List<ProjectWorkingRange>> _cache = new Dictionary<DateTime, List<ProjectWorkingRange>>();
    internal ProjectCalendarMath(IEnumerable<ProjectCalendar> calendars, int maxDays, CancellationToken token) {
        if (maxDays < 1 || maxDays > 366000) throw new ArgumentOutOfRangeException(nameof(maxDays));
        _calendars = calendars.Distinct().ToArray(); _maxDays = maxDays; _token = token;
        if (_calendars.Length == 0) throw new InvalidOperationException("Calculation requires an explicit calendar.");
        foreach (var calendar in _calendars) {
            calendar.Document.EnsureNotDisposed();
            if (!calendar.Attached) throw new InvalidOperationException("A calculation calendar has been removed.");
            ValidateProfile(calendar);
        }
    }
    internal static void Local(DateTime date) {
        if (date.Kind != DateTimeKind.Unspecified) throw new ArgumentException("Project arithmetic requires DateTimeKind.Unspecified local wall-clock values.");
    }
    private void ValidateProfile(ProjectCalendar calendar) {
        var seen = new HashSet<ProjectCalendar>();
        for (var current = calendar; current != null; current = current.BaseCalendar) {
            _token.ThrowIfCancellationRequested();
            if (!seen.Add(current)) throw new InvalidDataException("Calendar inheritance contains a cycle.");
            foreach (var exception in current.Exceptions) {
                var source = current.Document.Source?.Element(exception);
                int? type = (int?)source?.Element(source.Name.Namespace + "Type");
                if (type.HasValue && type != 1) throw new NotSupportedException("Recurring calendar exceptions require an expanded, qualified date range before calculation.");
            }
            if (current.HasUnqualifiedNativeRecurrence)
                throw new NotSupportedException("Unmodeled native calendar metadata must be qualified before calculation.");
        }
    }
    internal List<ProjectWorkingRange> Day(DateTime date) {
        _token.ThrowIfCancellationRequested(); Local(date); date = date.Date;
        if (_cache.TryGetValue(date, out var cached)) return cached;
        List<ProjectWorkingRange>? result = null;
        foreach (var calendar in _calendars) {
            var ranges = EffectiveDay(calendar, date);
            result = result == null ? ranges : Intersect(result, ranges);
        }
        result ??= new List<ProjectWorkingRange>();
        if (_cache.Count < _maxDays) _cache.Add(date, result);
        return result;
    }
    private List<ProjectWorkingRange> EffectiveDay(ProjectCalendar calendar, DateTime date) {
        if (date == DateTime.MaxValue.Date) throw new ArgumentOutOfRangeException(nameof(date), "The final DateTime date cannot contain a bounded full day.");
        var result = new List<ProjectWorkingRange>();
        if (date > DateTime.MinValue.Date) Append(calendar, date.AddDays(-1), date, date.AddDays(1), result);
        Append(calendar, date, date, date.AddDays(1), result);
        return Merge(result);
    }
    private void Append(ProjectCalendar calendar, DateTime ownerDate, DateTime from, DateTime to, List<ProjectWorkingRange> ranges) {
        var pattern = Resolve(calendar, ownerDate);
        foreach (var interval in pattern) {
            _token.ThrowIfCancellationRequested();
            if (!interval.From.HasValue || !interval.To.HasValue || interval.From < TimeSpan.Zero || interval.From >= TimeSpan.FromDays(1) ||
                interval.To < TimeSpan.Zero || interval.To >= TimeSpan.FromDays(1) || interval.From == interval.To)
                throw new InvalidDataException("A working interval must have distinct, valid local start and finish times.");
            DateTime start = ownerDate.Add(interval.From.Value), finish = ownerDate.Add(interval.To.Value);
            if (finish <= start) finish = finish.AddDays(1);
            if (start < from) start = from; if (finish > to) finish = to;
            if (finish > start) ranges.Add(new ProjectWorkingRange(start, finish));
        }
    }
    private IEnumerable<ProjectWorkingInterval> Resolve(ProjectCalendar calendar, DateTime date) {
        // Exceptions inherited from a base remain effective even when a derived calendar supplies an ordinary weekday.
        for (var current = calendar; current != null; current = current.BaseCalendar) {
            _token.ThrowIfCancellationRequested();
            var exceptions = current.Exceptions.Where(e => Covers(e.FromDate, e.ToDate, date)).ToArray();
            var legacy = current.WeekDays.Where(d => d.Day == null && Covers(d.FromDate, d.ToDate, date)).ToArray();
            if (exceptions.Length + legacy.Length > 1) throw new InvalidDataException("Overlapping calendar exceptions are ambiguous.");
            if (exceptions.Length == 1) return Times(exceptions[0].IsWorking, exceptions[0].WorkingTimes);
            if (legacy.Length == 1) return Times(legacy[0].IsWorking, legacy[0].WorkingTimes);
        }
        for (var current = calendar; current != null; current = current.BaseCalendar) {
            _token.ThrowIfCancellationRequested();
            var weeks = current.WorkWeeks.Where(w => Covers(w.FromDate, w.ToDate, date)).ToArray();
            if (weeks.Length > 1) throw new InvalidDataException("Overlapping work weeks are ambiguous.");
            var days = (weeks.Length == 1 ? weeks[0].WeekDays : Enumerable.Empty<ProjectWeekDay>()).Where(d => d.Day == date.DayOfWeek).ToArray();
            if (days.Length == 0) days = current.WeekDays.Where(d => d.Day == date.DayOfWeek).ToArray();
            if (days.Length > 1) throw new InvalidDataException("Duplicate weekday declarations are ambiguous.");
            if (days.Length == 1 && days[0].IsWorking.HasValue) return Times(days[0].IsWorking, days[0].WorkingTimes);
        }
        throw new InvalidDataException("The calendar does not define or inherit " + date.DayOfWeek + ".");
    }
    private static bool Covers(DateTime? from, DateTime? to, DateTime date) {
        if (from.HasValue) Local(from.Value); if (to.HasValue) Local(to.Value);
        if (from?.Date > to?.Date) throw new InvalidDataException("A calendar date range is reversed.");
        return (!from.HasValue || from.Value.Date <= date) && (!to.HasValue || to.Value.Date >= date);
    }
    private static IEnumerable<ProjectWorkingInterval> Times(bool? working, ProjectCollection<ProjectWorkingInterval> intervals) {
        if (!working.HasValue) throw new InvalidDataException("Calendar working status is unspecified.");
        if (working == false) return Array.Empty<ProjectWorkingInterval>();
        if (intervals.Count == 0) throw new InvalidDataException("A working calendar day has no intervals.");
        return intervals;
    }
    internal DateTime Snap(DateTime value, bool forward) {
        Local(value);
        DateTime day = value.Date;
        for (int count = 0; count < _maxDays; count++, day = day.AddDays(forward ? 1 : -1)) {
            var ranges = Day(day);
            foreach (var range in forward ? ranges : ranges.AsEnumerable().Reverse()) {
                if (forward && value < range.Finish) return value <= range.Start ? range.Start : value;
                if (!forward && value > range.Start) return value >= range.Finish ? range.Finish : value;
            }
        }
        throw new InvalidOperationException("The calendar search exceeded MaxCalendarDays without reaching working time.");
    }
    internal DateTime Add(DateTime start, decimal minutes) {
        Local(start); _token.ThrowIfCancellationRequested();
        if (minutes == 0) return start;
        bool forward = minutes > 0;
        long remaining = checked((long)decimal.Round(Math.Abs(minutes) * TimeSpan.TicksPerMinute, 0, MidpointRounding.AwayFromZero));
        DateTime day = start.Date, position = start;
        for (int count = 0; count < _maxDays; count++, day = day.AddDays(forward ? 1 : -1)) {
            var ranges = Day(day);
            foreach (var range in forward ? ranges : ranges.AsEnumerable().Reverse()) {
                DateTime edge = forward ? (position > range.Start ? position : range.Start) : (position < range.Finish ? position : range.Finish);
                long available = forward ? range.Finish.Ticks - edge.Ticks : edge.Ticks - range.Start.Ticks;
                if (available <= 0) continue;
                if (remaining <= available) return edge.AddTicks(forward ? remaining : -remaining);
                remaining -= available;
            }
        }
        throw new InvalidOperationException("Working-time arithmetic exceeded MaxCalendarDays.");
    }
    internal decimal Between(DateTime start, DateTime finish) {
        Local(start); Local(finish); _token.ThrowIfCancellationRequested();
        if (finish < start) return -Between(finish, start);
        if ((finish.Date - start.Date).TotalDays >= _maxDays) throw new InvalidOperationException("Working-time range exceeds MaxCalendarDays.");
        long ticks = 0;
        for (DateTime day = start.Date; day <= finish.Date; day = day.AddDays(1)) {
            foreach (var range in Day(day)) {
                long from = Math.Max(start.Ticks, range.Start.Ticks), to = Math.Min(finish.Ticks, range.Finish.Ticks);
                if (to > from) ticks = checked(ticks + to - from);
            }
        }
        return ticks / (decimal)TimeSpan.TicksPerMinute;
    }
    internal static List<ProjectWorkingRange> Merge(List<ProjectWorkingRange> ranges) {
        ranges.Sort((a, b) => a.Start.CompareTo(b.Start));
        var result = new List<ProjectWorkingRange>();
        foreach (var range in ranges) {
            if (result.Count == 0 || range.Start > result[result.Count - 1].Finish) result.Add(range);
            else {
                var last = result[result.Count - 1];
                if (range.Finish > last.Finish) result[result.Count - 1] = new ProjectWorkingRange(last.Start, range.Finish);
            }
        }
        return result;
    }
    private static List<ProjectWorkingRange> Intersect(List<ProjectWorkingRange> first, List<ProjectWorkingRange> second) {
        var result = new List<ProjectWorkingRange>(); int left = 0, right = 0;
        while (left < first.Count && right < second.Count) {
            DateTime start = first[left].Start > second[right].Start ? first[left].Start : second[right].Start;
            DateTime finish = first[left].Finish < second[right].Finish ? first[left].Finish : second[right].Finish;
            if (finish > start) result.Add(new ProjectWorkingRange(start, finish));
            if (first[left].Finish <= second[right].Finish) left++; else right++;
        }
        return result;
    }
}
