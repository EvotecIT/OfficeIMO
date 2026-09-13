namespace OfficeIMO.Project;

/// <summary>Entity collection represented by a mapped table.</summary>
public enum ProjectDataKind {
    /// <summary>Task identities, outline, dates and quantities.</summary>
    Tasks,
    /// <summary>Resource identities, kinds and allocation limits.</summary>
    Resources,
    /// <summary>Task/resource assignments.</summary>
    Assignments,
    /// <summary>Calendar identities, inheritance and ordinary working weeks.</summary>
    Calendars
}

/// <summary>Portable columns with fixed units; dates are local ISO 8601 values without timezone conversion.</summary>
public enum ProjectDataField {
    /// <summary>Stable nonnegative integer entity UID.</summary>
    Uid,
    /// <summary>Entity name.</summary>
    Name,
    /// <summary>Parent task UID; blank means a root task.</summary>
    ParentUid,
    /// <summary>Explicit summary marker, true or false.</summary>
    Summary,
    /// <summary>Calendar UID.</summary>
    CalendarUid,
    /// <summary>Stored local start.</summary>
    Start,
    /// <summary>Stored local finish.</summary>
    Finish,
    /// <summary>Duration in minutes.</summary>
    DurationMinutes,
    /// <summary>Whether duration minutes are elapsed rather than working.</summary>
    DurationElapsed,
    /// <summary>Work in minutes; material assignments use the native quantity encoding.</summary>
    WorkMinutes,
    /// <summary>Stored cost in project currency.</summary>
    Cost,
    /// <summary>Duration completion percentage from zero to one hundred.</summary>
    PercentComplete,
    /// <summary>Resource kind: Work, Material or Cost.</summary>
    ResourceType,
    /// <summary>Maximum allocation as a fraction; one means one hundred percent.</summary>
    MaxUnits,
    /// <summary>Assignment task UID.</summary>
    TaskUid,
    /// <summary>Assignment resource UID.</summary>
    ResourceUid,
    /// <summary>Assignment units as a fraction for labor or a quantity for materials.</summary>
    Units,
    /// <summary>Inherited calendar UID.</summary>
    BaseCalendarUid,
    /// <summary>Sunday intervals, comma-separated HH:mm:ss-HH:mm:ss; '-' means closed, blank means inherited.</summary>
    Sunday,
    /// <summary>Monday ordinary working intervals.</summary>
    Monday,
    /// <summary>Tuesday ordinary working intervals.</summary>
    Tuesday,
    /// <summary>Wednesday ordinary working intervals.</summary>
    Wednesday,
    /// <summary>Thursday ordinary working intervals.</summary>
    Thursday,
    /// <summary>Friday ordinary working intervals.</summary>
    Friday,
    /// <summary>Saturday ordinary working intervals.</summary>
    Saturday
}

/// <summary>A typed field bound to an external column header.</summary>
public sealed class ProjectDataColumn {
    /// <summary>Creates an explicit mapping; header comparison is ordinal and case-sensitive.</summary>
    public ProjectDataColumn(ProjectDataField field, string header) {
        if (!Enum.IsDefined(typeof(ProjectDataField), field)) throw new ArgumentOutOfRangeException(nameof(field));
        if (string.IsNullOrWhiteSpace(header)) throw new ArgumentException("A column header is required.", nameof(header));
        Field = field; Header = header;
    }
    /// <summary>Project field.</summary>
    public ProjectDataField Field { get; }
    /// <summary>External column header.</summary>
    public string Header { get; }
}

/// <summary>Bounded immutable text table shared by CSV and Excel adapters.</summary>
public sealed class ProjectDataTable {
    /// <summary>Copies headers and rows, rejecting inconsistent widths and enforcing row/cell bounds.</summary>
    public ProjectDataTable(IEnumerable<string> headers, IEnumerable<IReadOnlyList<string?>> rows, int maxRows = 100000,
        int maxCells = 2000000, int maxCellCharacters = 32767) {
        if (headers == null) throw new ArgumentNullException(nameof(headers));
        if (rows == null) throw new ArgumentNullException(nameof(rows));
        if (maxRows < 1 || maxCells < 1 || maxCellCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maxRows));
        var names = headers.Take(65).ToArray();
        if (names.Length == 0 || names.Length > 64 || names.Any(string.IsNullOrWhiteSpace) || names.Distinct(StringComparer.Ordinal).Count() != names.Length
            || names.Any(n => n.Length > maxCellCharacters)) throw new ArgumentException("Headers must be nonempty, distinct and bounded.", nameof(headers));
        var copy = new List<IReadOnlyList<string?>>();
        foreach (var row in rows) {
            if (copy.Count >= maxRows || (long)(copy.Count + 1) * names.Length > maxCells) throw new InvalidDataException("Mapped table exceeds its row or cell limit.");
            if (row == null || row.Count != names.Length) throw new InvalidDataException("Mapped row width differs from its headers.");
            var values = row.ToArray();
            if (values.Any(v => v?.Length > maxCellCharacters)) throw new InvalidDataException("Mapped cell exceeds its character limit.");
            copy.Add(Array.AsReadOnly(values));
        }
        Headers = Array.AsReadOnly(names); Rows = copy.AsReadOnly();
    }
    /// <summary>External headers.</summary>
    public IReadOnlyList<string> Headers { get; }
    /// <summary>Rows in source order. Null and empty cells both import as absent values.</summary>
    public IReadOnlyList<IReadOnlyList<string?>> Rows { get; }
}

/// <summary>A table and explicit field map for one entity collection.</summary>
public sealed class ProjectMappedTable {
    /// <summary>Creates a bounded field map. Invalid entity/field combinations are rejected during import.</summary>
    public ProjectMappedTable(ProjectDataKind kind, ProjectDataTable table, IEnumerable<ProjectDataColumn> columns) {
        if (!Enum.IsDefined(typeof(ProjectDataKind), kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        Kind = kind; Table = table ?? throw new ArgumentNullException(nameof(table));
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        var copy = columns.Take(65).ToArray();
        if (copy.Length == 0 || copy.Length > 64 || copy.Any(c => c == null) || copy.Select(c => c.Field).Distinct().Count() != copy.Length
            || copy.Select(c => c.Header).Distinct(StringComparer.Ordinal).Count() != copy.Length) throw new ArgumentException("Mappings must be distinct and bounded.", nameof(columns));
        Columns = Array.AsReadOnly(copy);
    }
    /// <summary>Entity collection.</summary>
    public ProjectDataKind Kind { get; }
    /// <summary>External table.</summary>
    public ProjectDataTable Table { get; }
    /// <summary>Field mappings.</summary>
    public IReadOnlyList<ProjectDataColumn> Columns { get; }
}

/// <summary>Import policy for a new document. Existing documents are never modified.</summary>
public sealed class ProjectDataImportOptions {
    /// <summary>Project name.</summary>
    public string? Name { get; set; }
    /// <summary>Project start date; no schedule calculation is implied.</summary>
    public DateTime? Start { get; set; }
    /// <summary>Project calendar UID, if included in the calendar table.</summary>
    public int? CalendarUid { get; set; }
    /// <summary>Permit unmapped columns and return their names in the loss report. False rejects them.</summary>
    public bool AllowUnmappedColumns { get; set; }
    /// <summary>Maximum rows across all input tables.</summary>
    public int MaxRows { get; set; } = 100000;
    /// <summary>Maximum outline or calendar inheritance depth.</summary>
    public int MaxDepth { get; set; } = 100;
}

/// <summary>New imported document and explicit ignored-column notices.</summary>
public sealed class ProjectDataImportResult {
    internal ProjectDataImportResult(ProjectDocument document, string[] notices) { Document = document; Notices = Array.AsReadOnly(notices); }
    /// <summary>Caller-owned document; dispose it when finished.</summary>
    public ProjectDocument Document { get; }
    /// <summary>Ignored input columns. Identity conflicts, malformed values and dangling references always reject the whole import.</summary>
    public IReadOnlyList<string> Notices { get; }
}

/// <summary>Mapped projection with an explicit description of its fidelity boundary.</summary>
public sealed class ProjectDataExportResult {
    internal ProjectDataExportResult(ProjectMappedTable[] tables, string[] notices) { Tables = Array.AsReadOnly(tables); Notices = Array.AsReadOnly(notices); }
    /// <summary>Four entity tables with explicit field maps.</summary>
    public IReadOnlyList<ProjectMappedTable> Tables { get; }
    /// <summary>Projection omissions. A table export is not a full-fidelity replacement for Project XML.</summary>
    public IReadOnlyList<string> Notices { get; }

    /// <summary>Returns a new projection with selected, reordered or renamed columns for one table. Omitted fields are added to the loss notices; imports still enforce their required mappings.</summary>
    public ProjectDataExportResult WithColumns(ProjectDataKind kind, IEnumerable<ProjectDataColumn> columns) {
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        var selected = columns.Take(65).ToArray();
        if (selected.Length == 0 || selected.Length > 64 || selected.Any(c => c == null)
            || selected.Select(c => c.Field).Distinct().Count() != selected.Length || selected.Select(c => c.Header).Distinct(StringComparer.Ordinal).Count() != selected.Length)
            throw new ArgumentException("Selected columns must be distinct and bounded.", nameof(columns));
        var source = Tables.FirstOrDefault(t => t.Kind == kind) ?? throw new ArgumentOutOfRangeException(nameof(kind));
        var indexes = source.Columns.Select((c, i) => (c.Field, Index: i)).ToDictionary(c => c.Field, c => c.Index);
        if (selected.Any(c => !indexes.ContainsKey(c.Field))) throw new ArgumentException("A selected field is not present in this projection.", nameof(columns));
        var table = new ProjectDataTable(selected.Select(c => c.Header), source.Table.Rows.Select(r => selected.Select(c => r[indexes[c.Field]]).ToArray()),
            Math.Max(1, source.Table.Rows.Count), Math.Max(1, checked(source.Table.Rows.Count * selected.Length)));
        var replacement = new ProjectMappedTable(kind, table, selected);
        var omitted = source.Columns.Where(c => !selected.Any(s => s.Field == c.Field)).Select(c => c.Field.ToString()).ToArray();
        return new ProjectDataExportResult(Tables.Select(t => t.Kind == kind ? replacement : t).ToArray(), omitted.Length == 0 ? Notices.ToArray()
            : Notices.Concat(new[] { kind + ": omitted mapped fields " + string.Join(", ", omitted) }).ToArray());
    }
}
