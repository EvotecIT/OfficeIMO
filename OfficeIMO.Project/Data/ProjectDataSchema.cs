using System.Globalization;

namespace OfficeIMO.Project;

internal static class ProjectDataSchema {
    internal static readonly ProjectDataField[] TaskFields = { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.ParentUid,
        ProjectDataField.Summary, ProjectDataField.CalendarUid, ProjectDataField.Start, ProjectDataField.Finish,
        ProjectDataField.DurationMinutes, ProjectDataField.DurationElapsed, ProjectDataField.WorkMinutes, ProjectDataField.Cost, ProjectDataField.PercentComplete };
    internal static readonly ProjectDataField[] ResourceFields = { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.ResourceType, ProjectDataField.CalendarUid, ProjectDataField.MaxUnits };
    internal static readonly ProjectDataField[] AssignmentFields = { ProjectDataField.Uid, ProjectDataField.TaskUid, ProjectDataField.ResourceUid, ProjectDataField.Units,
        ProjectDataField.Start, ProjectDataField.Finish, ProjectDataField.WorkMinutes, ProjectDataField.Cost };
    internal static readonly ProjectDataField[] CalendarFields = { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.BaseCalendarUid,
        ProjectDataField.Sunday, ProjectDataField.Monday, ProjectDataField.Tuesday, ProjectDataField.Wednesday, ProjectDataField.Thursday, ProjectDataField.Friday, ProjectDataField.Saturday };
    internal static ProjectDataField[] Fields(ProjectDataKind kind) => kind switch {
        ProjectDataKind.Tasks => TaskFields, ProjectDataKind.Resources => ResourceFields,
        ProjectDataKind.Assignments => AssignmentFields, ProjectDataKind.Calendars => CalendarFields,
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };
    internal static string? Text(object? value) => value switch {
        null => null, DateTime date when date.Kind == DateTimeKind.Unspecified => date.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture),
        DateTime => throw new InvalidDataException("Project table dates must use DateTimeKind.Unspecified; timezone-bearing values cannot be exported implicitly."),
        bool flag => flag ? "true" : "false", IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString()
    };
    internal static ProjectMappedTable Export(ProjectDataKind kind, IEnumerable<string?[]> rows, int maxRows, int maxCells) {
        var fields = Fields(kind);
        return new ProjectMappedTable(kind, new ProjectDataTable(fields.Select(f => f.ToString()), rows, maxRows, maxCells), fields.Select(f => new ProjectDataColumn(f, f.ToString())));
    }
    internal static string? WeekDay(ProjectCalendar calendar, DayOfWeek day) {
        var item = calendar.WeekDays.FirstOrDefault(d => d.Day == day);
        if (item == null) return null;
        if (item.IsWorking == false) return "-";
        if (item.IsWorking != true || item.WorkingTimes.Count == 0 || item.WorkingTimes.Any(t => !t.From.HasValue || !t.To.HasValue))
            throw new InvalidDataException("Calendar " + calendar.Uid + " has an incomplete ordinary weekday declaration.");
        return string.Join(",", item.WorkingTimes.Select(t => t.From!.Value.ToString("c", CultureInfo.InvariantCulture) + "-" + t.To!.Value.ToString("c", CultureInfo.InvariantCulture)));
    }

    internal sealed class Row {
        private readonly ProjectMappedTable _table;
        private readonly Dictionary<ProjectDataField, int> _indexes;
        private readonly IReadOnlyList<string?> _values;
        internal Row(ProjectMappedTable table, Dictionary<ProjectDataField, int> indexes, IReadOnlyList<string?> values, int number) {
            _table = table; _indexes = indexes; _values = values; Number = number;
        }
        internal int Number { get; }
        internal string? Get(ProjectDataField field) => _indexes.TryGetValue(field, out int index) && !string.IsNullOrEmpty(_values[index]) ? _values[index] : null;
        internal InvalidDataException Error(ProjectDataField field, string reason) => new InvalidDataException(_table.Kind + " row " + Number + ", " + field + ": " + reason);
        internal int? Integer(ProjectDataField field, bool required = false, bool allowUnassigned = false) {
            var text = Get(field);
            if (text == null) { if (required) throw Error(field, "a value is required"); return null; }
            if (allowUnassigned && text == "-1") return -1;
            if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value))
                throw Error(field, allowUnassigned ? "expected a nonnegative integer or -1 for unassigned" : "expected a nonnegative integer");
            return value;
        }
        internal decimal? Decimal(ProjectDataField field) {
            var text = Get(field); if (text == null) return null;
            if (!OfficeInvariantDecimal.TryParseExact(text, false, out decimal value) || value < 0)
                throw Error(field, "expected an exactly representable nonnegative invariant decimal");
            return value;
        }
        internal bool? Flag(ProjectDataField field) {
            var text = Get(field); if (text == null) return null;
            if (!bool.TryParse(text, out bool value)) throw Error(field, "expected true or false"); return value;
        }
        internal DateTime? Date(ProjectDataField field) {
            var text = Get(field); if (text == null) return null;
            if (!DateTime.TryParseExact(text, new[] { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.FFFFFFF" },
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var value)) throw Error(field, "expected a local ISO date without timezone suffix");
            return DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        }
        internal string? Name() => _indexes.TryGetValue(ProjectDataField.Name, out int index) ? _values[index] : null;
    }

    internal static Row[] Read(ProjectMappedTable table, ProjectDataImportOptions options, List<string> notices) {
        var supported = new HashSet<ProjectDataField>(Fields(table.Kind));
        var indexes = new Dictionary<ProjectDataField, int>();
        foreach (var column in table.Columns) {
            if (!supported.Contains(column.Field)) throw new InvalidDataException(column.Field + " cannot be mapped in " + table.Kind + ".");
            int index = -1;
            for (int i = 0; i < table.Table.Headers.Count; i++) if (table.Table.Headers[i] == column.Header) { index = i; break; }
            if (index < 0) throw new InvalidDataException("Missing mapped header: " + column.Header);
            indexes.Add(column.Field, index);
        }
        foreach (var required in table.Kind == ProjectDataKind.Assignments
            ? new[] { ProjectDataField.Uid, ProjectDataField.TaskUid, ProjectDataField.ResourceUid }
            : new[] { ProjectDataField.Uid, ProjectDataField.Name })
            if (!indexes.ContainsKey(required)) throw new InvalidDataException(table.Kind + " requires a mapping for " + required + ".");
        for (int i = 0; i < table.Table.Headers.Count; i++) if (!indexes.ContainsValue(i)) {
            if (!options.AllowUnmappedColumns) throw new InvalidDataException("Unmapped input column: " + table.Table.Headers[i]);
            notices.Add(table.Kind + ": ignored column " + table.Table.Headers[i]);
        }
        return table.Table.Rows.Select((r, i) => new Row(table, indexes, r, i + 1)).ToArray();
    }
}
