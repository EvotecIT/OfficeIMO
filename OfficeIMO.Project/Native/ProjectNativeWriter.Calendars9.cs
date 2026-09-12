namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private bool CalendarProjectionChanged(ProjectCalendar calendar) {
        // A removed work week or replaced ancestor must invalidate exceptions generated on a previous save.
        for (var current = calendar; current != null; current = current.BaseCalendar) {
            string path = Path(current, "Calendar");
            if (ChangedTree(path + "/Week") || Changed(path + "/BaseCalendar")) return true;
        }
        return false;
    }
    private static bool HasWorkWeeks(ProjectCalendar calendar) {
        for (var current = calendar; current != null; current = current.BaseCalendar)
            if (current.WorkWeeks.Count != 0) return true;
        return false;
    }

    private byte[] CalendarPattern9(ProjectCalendar calendar) {
        var dates = new SortedSet<DateTime>();
        void AddRange(DateTime? from, DateTime? to) {
            if (!from.HasValue || !to.HasValue) throw new NotSupportedException("MPP9 calendar overrides require a finite date range.");
            _ = CalendarDay(from); _ = CalendarDay(to);
            if (to.Value.Date < from.Value.Date) throw new ArgumentException("Calendar override dates are reversed.");
            int count = checked((int)(to.Value.Date - from.Value.Date).TotalDays + 1);
            if (count > ushort.MaxValue) throw new NotSupportedException("MPP9 calendar override expansion exceeds 65535 dates.");
            for (int i = 0; i < count; i++) {
                _token.ThrowIfCancellationRequested(); dates.Add(from.Value.Date.AddDays(i));
                if (dates.Count > ushort.MaxValue) throw new NotSupportedException("MPP9 calendar override expansion exceeds its record budget.");
                if (424L + dates.Count * 64L > _options.MaxOutputBytes)
                    throw OfficeIMO.Core.Internal.OfficeOutputLimit.Create("MPP9 calendar override expansion exceeds its byte budget.");
            }
        }
        foreach (var exception in calendar.Exceptions) AddRange(exception.FromDate, exception.ToDate);
        bool weeks = HasWorkWeeks(calendar);
        if (weeks) {
            // Materialize ancestor work weeks on derived calendars as well: their
            // ordinary weekday overrides must still precede a base's work week.
            for (var current = calendar; current != null; current = current.BaseCalendar)
                foreach (var week in current.WorkWeeks) AddRange(week.FromDate, week.ToDate);
            Loss("PROJECT_NATIVE_WORK_WEEK_FLATTENED", _profile.Generation + " represents bounded work weeks as dated calendar exceptions. The work-week structure and labels are omitted.", Path(calendar, "Calendar"));
        }
        if (calendar.Exceptions.Any(e => !string.IsNullOrEmpty(e.Name)))
            Loss("PROJECT_NATIVE_EXCEPTION_LABEL_LOSS", _profile.Generation + " calendar exceptions retain dates and working times but have no label field.", Path(calendar, "Calendar"));
        using var buffer = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(_options.MaxOutputBytes);
        using var writer = new BinaryWriter(buffer);
        writer.Write((ushort)dates.Count); writer.Write((ushort)0); writer.Write(CalendarWeek(calendar.WeekDays));
        var math = new ProjectCalendarMath(new[] { calendar }, ushort.MaxValue, _token);
        foreach (var date in dates) {
            _token.ThrowIfCancellationRequested(); var intervals = math.OwnerDayPattern(date);
            writer.Write(CalendarDay(date)); writer.Write(CalendarDay(date));
            writer.Write(CalendarDayPattern(intervals.Count != 0, intervals));
        }
        return buffer.ToArray();
    }
}
