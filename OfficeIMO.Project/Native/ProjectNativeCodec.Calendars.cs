namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    private static void ReadCalendars(ProjectDocument document, ProjectNativeTable table, Dictionary<uint, ProjectNativeValue> properties,
        ProjectLoadOptions options, CancellationToken token) {
        var bases = new Dictionary<int, int>();
        foreach (var record in table.Records) {
            token.ThrowIfCancellationRequested();
            record.Uid = record.Integer(0x0d400009) ?? throw new InvalidDataException("Native calendar UID missing.");
            if (record.Uid < 0) continue;
            CheckEntityBudget(document, options);
            int parent = record.Integer(0x0d400006) ?? 0;
            var calendar = new ProjectCalendar(document, record.Uid) { Name = record.Text(0x0d400001), IsBaseCalendar = record.Integer(0x0d400007) == -1 };
            calendar.Guid = EntityGuid(document, record, 0x0d40001b, 5);
            if (document.CalendarIndex.ContainsKey(calendar.Uid)) throw new InvalidDataException("Duplicate native calendar UID.");
            document.CalendarIndex.Add(calendar.Uid, calendar); document.Calendars.Items.Add(calendar); bases.Add(calendar.Uid, parent);
            var pattern = record.Value(0x0d400008);
            if (parent <= 0 && properties.TryGetValue(0x02401388, out var defaults)) {
                if (table.IsLegacy8) ReadCalendarPattern9(calendar, defaults, token);
                else ReadCalendarPattern(calendar, defaults, token);
            }
            if (pattern.HasValue) ReadCalendarPattern(calendar, pattern.Value, token);
            else if (table.IsLegacy8 && parent <= 0 && !properties.ContainsKey(0x02401388)) {
                if (calendar.Name != "Standard") throw new NotSupportedException("The Project 98 base calendar depends on an unavailable global calendar: " + calendar.Name);
                for (int day = 0; day < 7; day++) calendar.SetWorkingDay((DayOfWeek)day,
                    day == 0 || day == 6 ? Array.Empty<ProjectWorkingTime>() : new[] {
                        new ProjectWorkingTime(TimeSpan.FromHours(8), TimeSpan.FromHours(12)),
                        new ProjectWorkingTime(TimeSpan.FromHours(13), TimeSpan.FromHours(17)) });
                Warn(document, "PROJECT_NATIVE_IMPLICIT_STANDARD_CALENDAR", "The Project 98 Standard calendar omits its built-in weekly pattern. The qualified Monday-Friday 08:00-12:00 and 13:00-17:00 default is applied.", "/Calendar[UID=" + calendar.Uid + "]");
            }
        }
        foreach (var calendar in document.Calendars) {
            token.ThrowIfCancellationRequested();
            int parentId = bases[calendar.Uid];
            if (parentId > 0) {
                if (!document.CalendarIndex.TryGetValue(parentId, out var parent)) throw new InvalidDataException("Native base calendar is absent.");
                calendar.BindLoadedBaseCalendar(parent);
            }
        }
        ProjectXmlCodec.ValidateLoadedCalendarGraph(document, token);
        if (properties.TryGetValue(0x0240000e, out var projectCalendarName)) {
            var name = projectCalendarName.Unicode();
            document.Calendar = document.Calendars.FirstOrDefault(c => c.IsBaseCalendar == true && c.Name == name);
            if (document.Calendar == null) throw new InvalidDataException("Native project calendar name cannot be resolved.");
        }
    }
    private static void ReadCalendarPattern(ProjectCalendar calendar, ProjectNativeValue pattern, CancellationToken token) {
        if (calendar.Document.NativeInfo!.Profile.Version <= 9) { ReadCalendarPattern9(calendar, pattern, token, calendar.Document.NativeInfo.Profile.Version == 8); return; }
        if (pattern.Length < 420) throw new InvalidDataException("Truncated native calendar week.");
        for (int day = 0; day < 7; day++) {
            token.ThrowIfCancellationRequested();
            var data = pattern.Slice(day * 60, 60);
            if (data.UInt16() == 1) continue; // Explicit inheritance from the base calendar.
            if (data.UInt16() != 0) throw new NotSupportedException("Unqualified native working-day flags.");
            calendar.SetWorkingDay((DayOfWeek)day, ReadWorkingTimes(data));
        }
        if (pattern.Length == 420) return;
        int exceptions = pattern.Int32(420);
        if (exceptions < 0 || exceptions > (pattern.Length - 424) / 92) throw new InvalidDataException("Native calendar exception count is invalid.");
        int offset = 424;
        for (int index = 0; index < exceptions; index++) {
            token.ThrowIfCancellationRequested();
            var item = calendar.Exceptions.Add();
            item.FromDate = new DateTime(1983, 12, 31).AddDays(pattern.UInt16(offset));
            item.ToDate = new DateTime(1983, 12, 31).AddDays(pattern.UInt16(offset + 2) + 1).AddMinutes(-1);
            var times = ReadWorkingTimes(pattern.Slice(offset + 4, 60));
            item.IsWorking = times.Length != 0;
            foreach (var time in times) { var interval = item.WorkingTimes.Add(); interval.From = time.From; interval.To = time.To; }
            int labelLength = pattern.Int32(offset + 88);
            item.Name = pattern.Slice(checked(offset + 92), labelLength).Unicode();
            if (pattern.Int32(offset + 72) != 1) {
                calendar.HasUnqualifiedNativeRecurrence = true;
                Warn(calendar.Document, "PROJECT_NATIVE_CALENDAR_METADATA", "Recurring native calendar exceptions remain unqualified for calculation; the date span and source bytes are retained.", "/Calendar[UID=" + calendar.Uid + "]");
            }
            offset = checked((offset + 92 + labelLength + 3) & ~3);
        }
        if (offset == pattern.Length) return;
        int weeks = pattern.Int32(offset); offset += 4;
        if (weeks < 0 || weeks > (pattern.Length - offset) / 436) throw new InvalidDataException("Native work-week count is invalid.");
        for (int index = 0; index < weeks; index++) {
            token.ThrowIfCancellationRequested();
            var week = calendar.WorkWeeks.Add();
            for (int day = 0; day < 7; day++) {
                var data = pattern.Slice(offset + day * 60, 60);
                if (data.UInt16() == 1) continue;
                if (data.UInt16() != 0) throw new NotSupportedException("Unqualified native work-week day flags.");
                week.SetWorkingDay((DayOfWeek)day, ReadWorkingTimes(data));
            }
            week.FromDate = new DateTime(1983, 12, 31).AddDays(pattern.UInt16(offset + 420));
            week.ToDate = new DateTime(1983, 12, 31).AddDays(pattern.UInt16(offset + 422) + 1).AddMinutes(-1);
            int labelLength = pattern.Int32(offset + 432);
            week.Name = pattern.Slice(checked(offset + 436), labelLength).Unicode();
            offset = checked((offset + 436 + labelLength + 3) & ~3);
        }
        if (offset != pattern.Length) throw new NotSupportedException("Unqualified trailing native calendar metadata.");
    }
    private static ProjectWorkingTime[] ReadWorkingTimes(ProjectNativeValue day) {
        int count = day.UInt16(2);
        int maximum = day.Length == 40 ? 3 : 5, durationOffset = day.Length == 40 ? 16 : 20;
        if (count > maximum) throw new InvalidDataException("Native calendar shift count exceeds the generation's capacity.");
        var times = new ProjectWorkingTime[count];
        for (int i = 0; i < count; i++) {
            int from = day.UInt16(8 + i * 2), duration = day.Int32(durationOffset + i * 4);
            if (from >= 14400 || duration <= 0 || duration > 14400) throw new InvalidDataException("Native working interval is outside the clock range.");
            int finish = (from + duration) % 14400;
            times[i] = new ProjectWorkingTime(TimeSpan.FromTicks(from * (TimeSpan.TicksPerMinute / 10)), TimeSpan.FromTicks(finish * (TimeSpan.TicksPerMinute / 10)));
        }
        return times;
    }

    private static void ReadCalendarPattern9(ProjectCalendar calendar, ProjectNativeValue pattern, CancellationToken token, bool legacy8 = false) {
        int dayWidth = legacy8 ? 40 : 60, weekWidth = 4 + 7 * dayWidth, exceptionWidth = dayWidth + 4;
        if (pattern.Length < weekWidth) throw new InvalidDataException("Truncated legacy calendar week.");
        int count = pattern.UInt16();
        if (pattern.Length != weekWidth + count * exceptionWidth) throw new InvalidDataException("Legacy calendar exception records do not match their count.");
        for (int day = 0; day < 7; day++) {
            token.ThrowIfCancellationRequested(); var data = pattern.Slice(4 + day * dayWidth, dayWidth);
            if (data.UInt16() == 1) continue;
            if (data.UInt16() != 0) throw new NotSupportedException("Unqualified MPP9 working-day flags.");
            calendar.SetWorkingDay((DayOfWeek)day, ReadWorkingTimes(data));
        }
        for (int index = 0; index < count; index++) {
            token.ThrowIfCancellationRequested(); int offset = weekWidth + index * exceptionWidth;
            var item = calendar.Exceptions.Add();
            item.FromDate = new DateTime(1983, 12, 31).AddDays(pattern.UInt16(offset));
            item.ToDate = new DateTime(1983, 12, 31).AddDays(pattern.UInt16(offset + 2) + 1).AddMinutes(-1);
            var times = ReadWorkingTimes(pattern.Slice(offset + 4, dayWidth)); item.IsWorking = times.Length != 0;
            foreach (var time in times) { var interval = item.WorkingTimes.Add(); interval.From = time.From; interval.To = time.To; }
        }
    }
}
