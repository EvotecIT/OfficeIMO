using System.Globalization;

namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Imports explicitly mapped tables into a new document, preserving UIDs. Duplicate identities, missing required values, cycles, malformed values and dangling references reject the entire import.</summary>
    public static ProjectDataImportResult ImportTables(IEnumerable<ProjectMappedTable> tables, ProjectDataImportOptions? options = null, CancellationToken cancellationToken = default) {
        if (tables == null) throw new ArgumentNullException(nameof(tables));
        options ??= new ProjectDataImportOptions();
        var settings = new ProjectDataImportOptions { Name = options.Name, Start = options.Start, CalendarUid = options.CalendarUid,
            AllowUnmappedColumns = options.AllowUnmappedColumns, MaxRows = options.MaxRows, MaxDepth = options.MaxDepth };
        if (settings.MaxRows < 1 || settings.MaxDepth < 1 || settings.MaxDepth > 1000) throw new ArgumentOutOfRangeException(nameof(options));
        var input = tables.Take(5).ToArray();
        if (input.Length > 4 || input.Any(t => t == null) || input.Select(t => t.Kind).Distinct().Count() != input.Length)
            throw new ArgumentException("Supply at most one table per entity kind.", nameof(tables));
        if (input.Sum(t => (long)t.Table.Rows.Count) > settings.MaxRows) throw new InvalidDataException("Import exceeds MaxRows.");
        cancellationToken.ThrowIfCancellationRequested();
        var notices = new List<string>();
        var rows = input.ToDictionary(t => t.Kind, t => ProjectDataSchema.Read(t, settings, notices));
        ProjectDataSchema.Row[] Rows(ProjectDataKind kind) => rows.TryGetValue(kind, out var found) ? found : Array.Empty<ProjectDataSchema.Row>();
        var document = Create();
        try {
            using (document.BeginUpdate()) {
                document.Name = settings.Name; document.Settings.StartDate = settings.Start;
                document.ImportCalendars(Rows(ProjectDataKind.Calendars), settings.MaxDepth, cancellationToken);
                if (settings.CalendarUid.HasValue) document.Calendar = Lookup(document.CalendarIndex, settings.CalendarUid.Value, "project calendar");
                document.ImportTasks(Rows(ProjectDataKind.Tasks), settings.MaxDepth, cancellationToken);
                document.ImportResources(Rows(ProjectDataKind.Resources), cancellationToken);
                document.ImportAssignments(Rows(ProjectDataKind.Assignments), cancellationToken);
                document.Touch(true, true);
            }
            document.Validate(cancellationToken).ThrowIfErrors();
            return new ProjectDataImportResult(document, notices.ToArray());
        } catch { document.Dispose(); throw; }
    }

    private static T Lookup<T>(Dictionary<int, T> index, int uid, string role) => index.TryGetValue(uid, out var item)
        ? item : throw new InvalidDataException("Unknown " + role + " UID " + uid + ".");
    private static int ImportUid<T>(ProjectDataSchema.Row row, Dictionary<int, T> index) {
        int uid = row.Integer(ProjectDataField.Uid, true)!.Value;
        if (index.ContainsKey(uid)) throw row.Error(ProjectDataField.Uid, "duplicate identity " + uid);
        return uid;
    }
    private ProjectCalendar? ImportCalendar(ProjectDataSchema.Row row) => row.Integer(ProjectDataField.CalendarUid) is int uid
        ? Lookup(CalendarIndex, uid, "calendar") : null;

    private void ImportTasks(ProjectDataSchema.Row[] rows, int maxDepth, CancellationToken token) {
        var parents = new Dictionary<int, int?>();
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); int uid = ImportUid(row, TaskIndex);
            if (uid == 0 && row.Flag(ProjectDataField.Summary) != true)
                throw row.Error(ProjectDataField.Summary, "task UID zero is reserved for an explicit project summary");
            var task = new ProjectTask(this, uid) { Name = row.Name(), SourceSummary = row.Flag(ProjectDataField.Summary) ?? false,
                Calendar = ImportCalendar(row), Start = row.Date(ProjectDataField.Start), Finish = row.Date(ProjectDataField.Finish),
                Cost = row.Decimal(ProjectDataField.Cost), PercentComplete = row.Integer(ProjectDataField.PercentComplete) };
            if (row.Decimal(ProjectDataField.DurationMinutes) is decimal duration)
                task.Duration = new ProjectDuration(duration, ProjectDurationUnit.Minute, row.Flag(ProjectDataField.DurationElapsed) ?? false);
            else if (row.Get(ProjectDataField.DurationElapsed) != null) throw row.Error(ProjectDataField.DurationElapsed, "requires DurationMinutes");
            if (row.Decimal(ProjectDataField.WorkMinutes) is decimal work) task.Work = new ProjectWork(work);
            parents.Add(uid, row.Integer(ProjectDataField.ParentUid)); TaskIndex.Add(uid, task);
        }
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); int uid = row.Integer(ProjectDataField.Uid, true)!.Value; var task = TaskIndex[uid];
            var seen = new HashSet<int> { uid }; int? parentUid = parents[uid]; int depth = 0;
            for (var current = parentUid; current.HasValue; current = parents[current.Value]) {
                if (uid == 0 || current.Value == 0) throw row.Error(ProjectDataField.ParentUid, "the reserved project summary cannot be an outline parent or child");
                if (!parents.ContainsKey(current.Value)) throw row.Error(ProjectDataField.ParentUid, "unknown task " + current.Value);
                if (!seen.Add(current.Value) || ++depth > maxDepth) throw row.Error(ProjectDataField.ParentUid, "cycle or excessive outline depth");
            }
            task.Parent = parentUid.HasValue ? TaskIndex[parentUid.Value] : null;
            (task.Parent?.Children ?? Tasks).Items.Add(task);
        }
    }

    private void ImportResources(ProjectDataSchema.Row[] rows, CancellationToken token) {
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); int uid = ImportUid(row, ResourceIndex);
            ProjectResourceType? type = null;
            if (row.Get(ProjectDataField.ResourceType) is string text) {
                if (!Enum.TryParse<ProjectResourceType>(text, false, out var value) || !Enum.IsDefined(typeof(ProjectResourceType), value))
                    throw row.Error(ProjectDataField.ResourceType, "expected Work, Material or Cost");
                type = value;
            }
            var resource = new ProjectResource(this, uid) { Name = row.Name(), Type = type, Calendar = ImportCalendar(row) };
            if (row.Decimal(ProjectDataField.MaxUnits) is decimal units) resource.MaxUnits = ProjectUnits.Fraction(units);
            Resources.Items.Add(resource); ResourceIndex.Add(uid, resource);
        }
    }

    private void ImportAssignments(ProjectDataSchema.Row[] rows, CancellationToken token) {
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); int uid = ImportUid(row, AssignmentIndex);
            int taskUid = row.Integer(ProjectDataField.TaskUid, true)!.Value;
            int resourceUid = row.Integer(ProjectDataField.ResourceUid, required: true, allowUnassigned: true)!.Value;
            var task = Lookup(TaskIndex, taskUid, "assignment task");
            var resource = resourceUid == -1 ? null : Lookup(ResourceIndex, resourceUid, "assignment resource");
            if (resource != null && !AssignmentPairs.Add(PairKey(taskUid, resourceUid))) throw row.Error(ProjectDataField.ResourceUid, "duplicate task/resource assignment");
            var assignment = new ProjectAssignment(this, uid) { Task = task, Resource = resource, SourceTaskUid = taskUid, SourceResourceUid = resourceUid,
                Start = row.Date(ProjectDataField.Start), Finish = row.Date(ProjectDataField.Finish), Cost = row.Decimal(ProjectDataField.Cost) };
            if (row.Decimal(ProjectDataField.Units) is decimal units) assignment.Units = ProjectUnits.Fraction(units);
            if (row.Decimal(ProjectDataField.WorkMinutes) is decimal work) assignment.Work = new ProjectWork(work);
            Assignments.Items.Add(assignment); AssignmentIndex.Add(uid, assignment);
        }
    }

    private void ImportCalendars(ProjectDataSchema.Row[] rows, int maxDepth, CancellationToken token) {
        var parents = new Dictionary<int, int?>();
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); int uid = ImportUid(row, CalendarIndex);
            var calendar = new ProjectCalendar(this, uid) { Name = row.Name() };
            for (int day = 0; day < 7; day++) {
                var field = (ProjectDataField)((int)ProjectDataField.Sunday + day);
                string? text = row.Get(field); if (text == null) continue;
                var intervals = new List<ProjectWorkingTime>();
                if (text != "-") {
                    foreach (string range in text.Split(',')) {
                        if (intervals.Count >= 48) throw row.Error(field, "too many working intervals");
                        var pair = range.Split('-');
                        if (pair.Length != 2 || !TimeSpan.TryParseExact(pair[0], new[] { "c", @"hh\:mm" }, CultureInfo.InvariantCulture, out var from)
                            || !TimeSpan.TryParseExact(pair[1], new[] { "c", @"hh\:mm" }, CultureInfo.InvariantCulture, out var to)
                            || from < TimeSpan.Zero || from >= TimeSpan.FromDays(1) || to < TimeSpan.Zero || to >= TimeSpan.FromDays(1))
                            throw row.Error(field, "expected comma-separated HH:mm:ss-HH:mm:ss intervals");
                        intervals.Add(new ProjectWorkingTime(from, to));
                    }
                }
                calendar.SetWorkingDay((DayOfWeek)day, intervals.ToArray());
            }
            parents.Add(uid, row.Integer(ProjectDataField.BaseCalendarUid)); Calendars.Items.Add(calendar); CalendarIndex.Add(uid, calendar);
        }
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); int uid = row.Integer(ProjectDataField.Uid, true)!.Value;
            var seen = new HashSet<int> { uid }; int depth = 0;
            for (var current = parents[uid]; current.HasValue; current = parents[current.Value]) {
                if (!parents.ContainsKey(current.Value)) throw row.Error(ProjectDataField.BaseCalendarUid, "unknown calendar " + current.Value);
                if (!seen.Add(current.Value) || ++depth > maxDepth) throw row.Error(ProjectDataField.BaseCalendarUid, "cycle or excessive inheritance depth");
            }
            var calendar = CalendarIndex[uid];
            calendar.BaseCalendar = parents[uid].HasValue ? CalendarIndex[parents[uid]!.Value] : null;
            calendar.IsBaseCalendar = calendar.BaseCalendar == null;
        }
    }
}
