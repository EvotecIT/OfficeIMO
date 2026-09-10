namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private void WriteCalendars() {
        if (!_new && !ChangedTree("/Calendar") && !_calendarBindingsChanged && !_resourceGuidsChanged) return;
        using var editor = Editor("Cal", 0x16, 0x0d400009);
        var owners = _document.Resources.Where(r => r.Calendar != null).GroupBy(r => r.Calendar!).ToDictionary(g => g.Key, g => g.ToArray());
        if (_new) for (int i = 0; i < 4; i++) editor.AddReserved(i);
        Removed(editor, "Calendar", _document.CalendarIndex.Keys);
        foreach (var calendar in _document.Calendars) {
            _token.ThrowIfCancellationRequested(); string path = Path(calendar, "Calendar"); bool added = !editor.Contains(calendar.Uid);
            owners.TryGetValue(calendar, out var resources);
            var owner = resources?.Length == 1 ? resources[0] : null;
            if (calendar.IsBaseCalendar == false && owner == null) {
                AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_CALENDAR_OWNER", ProjectDiagnosticSeverity.Error,
                    "Each native derived calendar must belong to exactly one resource.", path)); continue;
            }
            if (added) { editor.Add(calendar.Uid); editor.Integer(calendar.Uid, 0x0d400009, calendar.Uid);
                Identity(editor, calendar.Uid, 0x0d40001b, EntityGuid(calendar, 5));
            }
            if (_profile.HasExtendedRecords && (added || Changed("/Settings/Calendar"))) editor.Set(calendar.Uid, 0x0d400022, new byte[] { calendar == _document.Calendar ? (byte)1 : (byte)0 });
            Handle(path + "/Uid");
            Field(editor, calendar.Uid, path, "Name", 0x0d400001, Kind.Text);
            Field(editor, calendar.Uid, path, "Guid", 0x0d40001b, Kind.Guid);
            Handle(path + "/BaseCalendar"); Handle(path + "/IsBaseCalendar");
            if (added || Changed(path + "/BaseCalendar")) editor.Integer(calendar.Uid, 0x0d400006, calendar.BaseCalendar?.Uid ?? (_profile.Version <= 9 ? -1 : 0));
            if (added || Changed(path + "/IsBaseCalendar") || _calendarBindingsChanged) {
                editor.Integer(calendar.Uid, 0x0d400007, calendar.IsBaseCalendar == false ? owner!.Uid : -1);
                editor.Set(calendar.Uid, 0x0d400000, new byte[] { calendar.IsBaseCalendar == false ? (byte)0 : (byte)1 });
            }
            if (added || Changed(path + "/BaseCalendar") || _calendarGuidsChanged) Identity(editor, calendar.Uid, 0x0d40001d, calendar.BaseCalendar == null ? Guid.Empty : EntityGuid(calendar.BaseCalendar, 5));
            if (added || _calendarBindingsChanged || _resourceGuidsChanged) Identity(editor, calendar.Uid, 0x0d40001c, owner == null ? Guid.Empty : EntityGuid(owner, 2));
            bool patternChanged = ChangedTree(path + "/Day") || ChangedTree(path + "/Exception") || ChangedTree(path + "/Week");
            if (!added && !patternChanged && !(_profile.Version <= 9 && (HasWorkWeeks(calendar) || CalendarProjectionChanged(calendar)))) continue;
            if (calendar.HasUnqualifiedNativeRecurrence) {
                AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_CALENDAR_RECURRENCE", ProjectDiagnosticSeverity.Error,
                    "This calendar contains unmodeled recurring exceptions. Its working pattern cannot be replaced safely.", path)); continue;
            }
            try {
                editor.Set(calendar.Uid, 0x0d400008, CalendarPattern(calendar));
                HandleTree(path + "/Day"); HandleTree(path + "/Exception"); HandleTree(path + "/Week");
            } catch (Exception ex) when (ex is ArgumentException || ex is OverflowException || ex is NotSupportedException) {
                AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_CALENDAR_VALUE", ProjectDiagnosticSeverity.Error, ex.Message, path));
            }
        }
        editor.Export(_replacements);
    }
    private byte[] CalendarPattern(ProjectCalendar calendar) {
        if (_profile.Version <= 9) return CalendarPattern9(calendar);
        using var buffer = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(_options.MaxOutputBytes);
        using var writer = new BinaryWriter(buffer);
        writer.Write(CalendarWeek(calendar.WeekDays)); writer.Write(calendar.Exceptions.Count);
        foreach (var exception in calendar.Exceptions) {
            _token.ThrowIfCancellationRequested(); var record = new byte[92];
            Put(record, 0, CalendarDay(exception.FromDate)); Put(record, 2, CalendarDay(exception.ToDate));
            Buffer.BlockCopy(CalendarDayPattern(exception.IsWorking, exception.WorkingTimes), 0, record, 4, 60);
            Put(record, 72, 1); byte[] label = Text(exception.Name ?? string.Empty); Put(record, 88, label.Length);
            writer.Write(record); writer.Write(label); Pad(writer);
        }
        writer.Write(calendar.WorkWeeks.Count);
        foreach (var week in calendar.WorkWeeks) {
            _token.ThrowIfCancellationRequested(); writer.Write(CalendarWeek(week.WeekDays));
            var record = new byte[16]; Put(record, 0, CalendarDay(week.FromDate)); Put(record, 2, CalendarDay(week.ToDate));
            byte[] label = Text(week.Name ?? string.Empty); Put(record, 12, label.Length);
            writer.Write(record); writer.Write(label); Pad(writer);
        }
        return buffer.ToArray();
    }
    private byte[] CalendarWeek(IEnumerable<ProjectWeekDay> days) {
        int width = _profile == ProjectNativeProfile.Mpp8 ? 40 : 60;
        var result = new byte[7 * width]; var listed = days.ToArray();
        if (listed.Any(d => !d.Day.HasValue || d.FromDate.HasValue || d.ToDate.HasValue))
            throw new NotSupportedException("Native ordinary weeks require weekday declarations; express date exceptions in Exceptions.");
        for (int day = 0; day < 7; day++) {
            var item = listed.SingleOrDefault(d => (int?)d.Day == day);
            if (item == null) result[day * width] = 1;
            else Buffer.BlockCopy(CalendarDayPattern(item.IsWorking, item.WorkingTimes), 0, result, day * width, width);
        }
        return result;
    }
    private byte[] CalendarDayPattern(bool? working, IEnumerable<ProjectWorkingInterval> intervals) {
        bool legacy8 = _profile == ProjectNativeProfile.Mpp8;
        int capacity = legacy8 ? 3 : 5, durationOffset = legacy8 ? 16 : 20, totalOffset = legacy8 ? 28 : 40;
        var result = new byte[legacy8 ? 40 : 60]; var times = intervals.ToArray();
        if (!working.HasValue) throw new ArgumentException("Native weekday working status must be explicit.");
        if (!working.Value && times.Length != 0) throw new ArgumentException("A non-working day cannot contain working intervals.");
        if (times.Length > capacity) throw new NotSupportedException("This native generation supports at most " + capacity + " working intervals per day.");
        Put(result, 2, checked((ushort)times.Length)); int total = 0;
        for (int i = 0; i < times.Length; i++) {
            if (!times[i].From.HasValue || !times[i].To.HasValue) throw new ArgumentException("Working interval endpoints are required.");
            int from = Exact(times[i].From!.Value.Ticks / (decimal)(TimeSpan.TicksPerMinute / 10));
            int to = Exact(times[i].To!.Value.Ticks / (decimal)(TimeSpan.TicksPerMinute / 10));
            int duration = to - from;
            if (duration <= 0) duration += 14400;
            Put(result, 8 + i * 2, checked((ushort)from)); Put(result, durationOffset + i * 4, duration); total += duration;
            Put(result, totalOffset + i * 4, total);
        }
        Put(result, 4, total); return result;
    }
    private static ushort CalendarDay(DateTime? date) {
        if (!date.HasValue) throw new NotSupportedException("Native exceptions and work weeks require explicit date bounds.");
        return checked((ushort)((date.Value.Date - new DateTime(1983, 12, 31)).Days));
    }
    private static void Put(byte[] bytes, int offset, int value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, offset, 4);
    private static void Put(byte[] bytes, int offset, ushort value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, offset, 2);
    private static void Pad(BinaryWriter writer) { while ((writer.BaseStream.Position & 3) != 0) writer.Write((byte)0); }
}
