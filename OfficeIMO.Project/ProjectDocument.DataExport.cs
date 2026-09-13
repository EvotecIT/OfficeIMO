namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Projects stored task/resource/assignment/calendar fields into four mapped tables. This deliberately omits dependencies, advanced scheduling data and native/source payloads; inspect Notices before using it as an interchange format.</summary>
    public ProjectDataExportResult ExportTables(bool allowLossyProjection = false, int maxRows = 100000, int maxCells = 2000000, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (!allowLossyProjection) throw new InvalidOperationException("Table export omits project settings and advanced data. Set allowLossyProjection explicitly and inspect the returned Notices, or save Project XML for a full model transfer.");
        if (HasActiveUpdate) throw new InvalidOperationException("Finish the update scope before exporting tables.");
        if (maxRows < 1 || maxCells < 1) throw new ArgumentOutOfRangeException(nameof(maxRows));
        if ((long)TaskIndex.Count + Resources.Count + Assignments.Count + Calendars.Count > maxRows) throw new InvalidOperationException("Table export exceeds MaxRows.");
        long cells = (long)TaskIndex.Count * ProjectDataSchema.TaskFields.Length + (long)Resources.Count * ProjectDataSchema.ResourceFields.Length
            + (long)Assignments.Count * ProjectDataSchema.AssignmentFields.Length + (long)Calendars.Count * ProjectDataSchema.CalendarFields.Length;
        if (cells > maxCells) throw new InvalidOperationException("Table export exceeds MaxCells.");
        long revision = Revision;
        string?[] Values(params object?[] values) { cancellationToken.ThrowIfCancellationRequested(); return values.Select(ProjectDataSchema.Text).ToArray(); }
        var tasks = ProjectDataSchema.Export(ProjectDataKind.Tasks, AllTasks.Select(t => Values(t.Uid, t.Name, t.Parent?.Uid, t.IsSummary, t.Calendar?.Uid,
            t.Start, t.Finish, t.Duration is ProjectDuration duration ? duration.Value * ProjectTimeUnits.MinutesPerUnit(duration.Unit, duration.IsElapsed, Settings) : (decimal?)null,
            t.Duration?.IsElapsed, t.Work?.Minutes, t.Cost, t.PercentComplete)), maxRows, maxCells);
        var resources = ProjectDataSchema.Export(ProjectDataKind.Resources, Resources.Select(r => Values(r.Uid, r.Name, r.Type, r.Calendar?.Uid, r.MaxUnits?.Value)), maxRows, maxCells);
        var assignments = ProjectDataSchema.Export(ProjectDataKind.Assignments, Assignments.Select(a => Values(a.Uid, a.Task?.Uid, a.Resource?.Uid ?? (a.SourceResourceUid < 0 ? -1 : a.SourceResourceUid), a.Units?.Value,
            a.Start, a.Finish, a.Work?.Minutes, a.Cost)), maxRows, maxCells);
        var calendars = ProjectDataSchema.Export(ProjectDataKind.Calendars, Calendars.Select(c => Values(c.Uid, c.Name, c.BaseCalendar?.Uid,
            ProjectDataSchema.WeekDay(c, DayOfWeek.Sunday), ProjectDataSchema.WeekDay(c, DayOfWeek.Monday), ProjectDataSchema.WeekDay(c, DayOfWeek.Tuesday),
            ProjectDataSchema.WeekDay(c, DayOfWeek.Wednesday), ProjectDataSchema.WeekDay(c, DayOfWeek.Thursday), ProjectDataSchema.WeekDay(c, DayOfWeek.Friday),
            ProjectDataSchema.WeekDay(c, DayOfWeek.Saturday))), maxRows, maxCells);
        if (Revision != revision) throw new InvalidOperationException("The project changed during table export.");
        cancellationToken.ThrowIfCancellationRequested();
        return new ProjectDataExportResult(new[] { tasks, resources, assignments, calendars }, new[] {
            "Only the fields listed in each table map are included. Project metadata/settings, dependencies, rates, progress detail, baselines, custom fields, exceptions, work weeks, external links and opaque source data are omitted.",
            "Durations are normalized to minutes; display units and estimate markers are omitted. Empty cells represent absent values. Supply project start and calendar UID explicitly when importing.",
            "Stored values are exported without recalculation. CSV/Excel tables are a bounded data projection, not a full-fidelity project backup."
        });
    }
}
