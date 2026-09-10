namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Validates identities, relationships, values, and preservation risk without calculating or writing.</summary>
    public ProjectReport Validate(CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        var diagnostics = new List<ProjectDiagnostic>();
        bool truncated = false;
        void Add(string code, string message, string location, ProjectDiagnosticSeverity severity = ProjectDiagnosticSeverity.Error, bool loss = false) {
            if (diagnostics.Count < 1000) diagnostics.Add(new ProjectDiagnostic(code, severity, message, location, loss));
            else truncated = true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (Settings.MinutesPerDay <= 0 || Settings.MinutesPerWeek <= 0 || Settings.DaysPerMonth <= 0)
            Add("PROJECT_WORKING_TIME", "Working-time conversion settings must be positive.", "/Project");
        if (Settings.ScheduleFromStart == true && !Settings.StartDate.HasValue)
            Add("PROJECT_START_REQUIRED", "Set a project start date when scheduling from the start.", "/Project/StartDate");
        if (Settings.ScheduleFromStart == false && !Settings.FinishDate.HasValue)
            Add("PROJECT_FINISH_REQUIRED", "Set a project finish date when scheduling from the finish.", "/Project/FinishDate");
        if (Settings.CurrencyDigits < 0 || Settings.CurrencyDigits > 10)
            Add("PROJECT_CURRENCY_DIGITS", "Currency precision must be between zero and ten.", "/Project/CurrencyDigits");
        if (Settings.CurrencyCode != null && Settings.CurrencyCode.Length != 3)
            Add("PROJECT_CURRENCY_CODE", "A currency code must contain three characters.", "/Project/CurrencyCode");
        if (Source == null && Settings.CurrencyCode == null)
            Add("PROJECT_CURRENCY_CODE", "New MSPDI documents require a currency code.", "/Project/CurrencyCode");
        CheckDate(Settings.StartDate, "/Project/StartDate", Add);
        CheckDate(Settings.FinishDate, "/Project/FinishDate", Add);
        CheckDate(Settings.StatusDate, "/Project/StatusDate", Add);
        CheckEnum(Settings.DefaultTaskType, "/Project/DefaultTaskType", Add);
        CheckClock(Settings.DefaultStartTime, "/Project/DefaultStartTime", Add);
        CheckClock(Settings.DefaultFinishTime, "/Project/DefaultFinishTime", Add);
        foreach (var task in AllTasks) {
            cancellationToken.ThrowIfCancellationRequested();
            string location = "/Task[UID=" + task.Uid + "]";
            CheckEnum(task.Type, location + "/Type", Add);
            CheckEnum(task.ConstraintType, location + "/ConstraintType", Add);
            if (task.ConstraintType.HasValue && (int)task.ConstraintType.Value >= 2 && !task.ConstraintDate.HasValue)
                Add("PROJECT_CONSTRAINT_DATE", "A date-bound constraint requires a constraint date.", location);
            if (task.Duration?.Value < 0 || task.ActualDuration?.Value < 0 || task.RemainingDuration?.Value < 0)
                Add("PROJECT_NEGATIVE_DURATION", "Task durations cannot be negative; use dependency lag for lead time.", location);
            CheckPercent(task.PercentComplete, location, Add); CheckPercent(task.PercentWorkComplete, location, Add); CheckPercent(task.PhysicalPercentComplete, location, Add);
            CheckDateRange(task.Start, task.Finish, location, Add);
            CheckDateRange(task.ActualStart, task.ActualFinish, location + "/Actual", Add);
            CheckDate(task.Deadline, location + "/Deadline", Add); CheckDate(task.ConstraintDate, location + "/ConstraintDate", Add);
            if (task.Calendar == null && task.SourceCalendarUid > 0)
                Add("PROJECT_CALENDAR_REFERENCE", "The task references a missing calendar.", location);
            CheckRich(task.Baselines, task.CustomFields, task.TimephasedData, location, Add, cancellationToken);
        }
        foreach (var resource in Resources) {
            cancellationToken.ThrowIfCancellationRequested();
            string location = "/Resource[UID=" + resource.Uid + "]";
            CheckEnum(resource.Type, location + "/Type", Add);
            if (resource.Calendar == null && resource.SourceCalendarUid > 0)
                Add("PROJECT_CALENDAR_REFERENCE", resource.Uid == 0 ? "The reserved resource row retains an implicit calendar reference not exported by Project." : "The resource references a missing calendar.", location,
                    resource.Uid == 0 ? ProjectDiagnosticSeverity.Warning : ProjectDiagnosticSeverity.Error);
            CheckRich(resource.Baselines, resource.CustomFields, resource.TimephasedData, location, Add, cancellationToken);
        }
        var assignmentPairs = new HashSet<long>();
        foreach (var assignment in Assignments) {
            cancellationToken.ThrowIfCancellationRequested();
            string location = "/Assignment[UID=" + assignment.Uid + "]";
            if (assignment.Task != null && assignment.Resource != null && !assignmentPairs.Add(PairKey(assignment.Task.Uid, assignment.Resource.Uid)))
                Add("PROJECT_ASSIGNMENT_DUPLICATE", "A task/resource assignment is duplicated.", location);
            if (assignment.Task == null || !assignment.Task.Attached)
                Add("PROJECT_TASK_REFERENCE", "The assignment references a missing task.", location);
            if (assignment.Resource == null && assignment.SourceResourceUid >= 0)
                Add("PROJECT_RESOURCE_REFERENCE", "The assignment references a missing resource.", location);
            CheckPercent(assignment.PercentWorkComplete, location, Add);
            if (assignment.Resource?.Type == ProjectResourceType.Cost && (assignment.Cost != null || assignment.ActualCost != null || assignment.RemainingCost != null))
                Add("PROJECT_COST_RESOURCE_IMPORT", "Stored cost-resource amounts are retained in XML, but Microsoft Project 2024 can discard them when importing even its own XML exports. Native application amount fidelity is not qualified.", location, ProjectDiagnosticSeverity.Warning);
            CheckDateRange(assignment.Start, assignment.Finish, location, Add);
            CheckDateRange(assignment.ActualStart, assignment.ActualFinish, location + "/Actual", Add);
            CheckRich(assignment.Baselines, assignment.CustomFields, assignment.TimephasedData, location, Add, cancellationToken);
        }
        if (Calendar == null && Settings.SourceCalendarUid > 0)
            Add("PROJECT_CALENDAR_REFERENCE", "The project references a missing calendar.", "/Project/CalendarUID");
        ValidateCalendars(Add, cancellationToken);
        ValidateDependencies(Add, cancellationToken);
        var fields = new HashSet<string>(StringComparer.Ordinal);
        foreach (var field in CustomFields) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(field.FieldId) || !fields.Add(field.FieldId!))
                Add("PROJECT_CUSTOM_FIELD_ID", "Custom-field definitions need unique nonempty field IDs.", "/Project/ExtendedAttributes");
            var lookupIds = new HashSet<int>();
            foreach (var value in field.LookupValues)
                if (!value.Id.HasValue || !lookupIds.Add(value.Id.Value))
                    Add("PROJECT_LOOKUP_ID", "Lookup values need unique IDs.", "/Project/ExtendedAttributes/" + field.FieldId);
        }
        if (IsScheduleStale) Add("PROJECT_SCHEDULE_STALE", "Schedule-affecting edits have been made; stored dates, work, and costs have not been recalculated.", "/Project", ProjectDiagnosticSeverity.Warning);
        if (StructureChanged && Source?.HasOpaqueStructures == true)
            Add("PROJECT_OPAQUE_REFERENCES", "Structural edits may invalidate references inside preserved, unmodeled XML. Review the source diagnostics and use Allow loss only when this risk is acceptable.", "/Project", ProjectDiagnosticSeverity.Warning, true);
        if (truncated) diagnostics.Add(new ProjectDiagnostic("PROJECT_VALIDATION_TRUNCATED", ProjectDiagnosticSeverity.Error, "Validation exceeded the diagnostic budget; additional findings are omitted.", "/Project"));
        return new ProjectReport(Revision, diagnostics);
    }

    private delegate void Finding(string code, string message, string location, ProjectDiagnosticSeverity severity = ProjectDiagnosticSeverity.Error, bool loss = false);
    private static void CheckEnum<T>(T? value, string location, Finding add) where T : struct, Enum {
        if (value.HasValue && !Enum.IsDefined(typeof(T), value.Value)) add("PROJECT_ENUM_VALUE", "The value is outside the supported enumeration.", location);
    }
    private static void CheckClock(TimeSpan? value, string location, Finding add) {
        if (value < TimeSpan.Zero || value >= TimeSpan.FromDays(1)) add("PROJECT_CLOCK_VALUE", "Clock values must fall within one day.", location);
    }
    private static void CheckDate(DateTime? date, string location, Finding add) {
        if (date.HasValue && date.Value.Kind != DateTimeKind.Unspecified)
            add("PROJECT_DATE_KIND", "Project dates must use DateTimeKind.Unspecified. Convert deliberately before assignment.", location);
    }
    private static void CheckDateRange(DateTime? start, DateTime? finish, string location, Finding add) {
        CheckDate(start, location + "/Start", add); CheckDate(finish, location + "/Finish", add);
        if (start.HasValue && finish.HasValue && start > finish)
            add("PROJECT_DATE_RANGE", "Start is later than finish.", location);
    }
    private static void CheckPercent(int? value, string location, Finding add) {
        if (value < 0 || value > 100) add("PROJECT_PERCENT", "A completion percentage must be between 0 and 100.", location);
    }
    private static void CheckRich(ProjectCollection<ProjectBaseline> baselines, ProjectCollection<ProjectCustomFieldValue> fields,
        ProjectCollection<ProjectTimephasedValue> timephased, string location, Finding add, CancellationToken token) {
        var numbers = new HashSet<int>();
        foreach (var baseline in baselines) {
            token.ThrowIfCancellationRequested();
            if (baseline.Owner is ProjectResource && (baseline.Start.HasValue || baseline.Finish.HasValue || baseline.TimephasedData.Count != 0))
                add("PROJECT_BASELINE_CONTEXT", "Resource baselines do not support dates or timephased data.", location + "/Baseline");
            if (!(baseline.Owner is ProjectTask) && (baseline.Duration.HasValue || baseline.FixedCost.HasValue))
                add("PROJECT_BASELINE_CONTEXT", "Only task baselines support duration and fixed cost.", location + "/Baseline");
            if (!baseline.Number.HasValue || baseline.Number < 0 || baseline.Number > 10 || !numbers.Add(baseline.Number.Value))
                add("PROJECT_BASELINE_NUMBER", "Baselines need a unique number from 0 through 10.", location);
            CheckDateRange(baseline.Start, baseline.Finish, location + "/Baseline", add);
            if (baseline.Duration?.Value < 0) add("PROJECT_NEGATIVE_DURATION", "A baseline duration cannot be negative.", location);
            CheckTimephased(baseline.TimephasedData, location + "/Baseline", add, token);
        }
        foreach (var field in fields) {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(field.FieldId)) add("PROJECT_CUSTOM_FIELD_ID", "A custom value requires a field ID.", location);
        }
        CheckTimephased(timephased, location, add, token);
    }
    private static void CheckTimephased(ProjectCollection<ProjectTimephasedValue> intervals, string location, Finding add, CancellationToken token) {
        foreach (var interval in intervals) {
            token.ThrowIfCancellationRequested();
            CheckDateRange(interval.Start, interval.Finish, location + "/TimephasedData", add);
        }
    }
}
