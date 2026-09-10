using System.Xml.Linq;

namespace OfficeIMO.Project;

internal sealed partial class ProjectScheduler {
    private void CheckStoredAssignmentDates(ProjectTaskSchedule[] tasks) {
        if (tasks.Length == 0) return;
        var results = tasks.ToDictionary(t => t.TaskUid);
        foreach (var assignment in _document.Assignments) {
            _token.ThrowIfCancellationRequested();
            if (assignment.Task == null || !results.TryGetValue(assignment.Task.Uid, out var result)) continue;
            if (assignment.Start.HasValue && assignment.Start != result.Start || assignment.Finish.HasValue && assignment.Finish != result.Finish)
                Error("PROJECT_ASSIGNMENT_DATE_RECALCULATION_REQUIRED",
                    "The proposed task dates differ from stored assignment dates. Applying them requires assignment/work/curve rescheduling, which the date-only profile cannot perform.", assignment.Task);
        }
    }
    private void CheckSourceProfile() {
        if (_document.NativeSource != null)
            _diagnostics.Add(new ProjectDiagnostic("PROJECT_NATIVE_SCHEDULE_PROJECTION", ProjectDiagnosticSeverity.Warning,
                "Calculation uses decoded scalar values and calendars. Native contours, splits, leveling, and other opaque scheduling records are not interpreted; this is a projection of the supported model.", "/Project"));
        foreach (var task in _document.AllTasks) {
            _token.ThrowIfCancellationRequested();
            var source = _document.Source?.Element(task);
            if (source == null) continue;
            if (Enabled(source, "ExternalTask") || Enabled(source, "IsSubproject") || Enabled(source, "Recurring") ||
                Enabled(source, "IgnoreResourceCalendar") || Nonzero(source, "LevelingDelay") || source.Element(source.Name.Namespace + "RecurringTask") != null)
                Error("PROJECT_TASK_SCHEDULING_PROFILE", "External/recurring tasks, leveling delays, and ignored resource calendars require additional scheduling semantics.", task);
        }
        foreach (var assignment in _document.Assignments) {
            _token.ThrowIfCancellationRequested();
            if (assignment.ActualStart.HasValue || assignment.ActualFinish.HasValue || assignment.ActualWork?.Minutes > 0 || assignment.PercentWorkComplete > 0)
                Error("PROJECT_PROGRESS_SCHEDULING", "Assignment progress requires interval-aware rescheduling.", assignment.Task);
            var source = _document.Source?.Element(assignment);
            if (source != null && (Nonzero(source, "Delay") || Nonzero(source, "WorkContour") || Nonzero(source, "LevelingDelay")))
                Error("PROJECT_ASSIGNMENT_SCHEDULING_PROFILE", "Assignment delays, non-flat work contours, and leveling delays require independent assignment scheduling.", assignment.Task);
        }
    }
    private static bool Enabled(XElement source, string name) {
        string? value = source.Element(source.Name.Namespace + name)?.Value;
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
    private static bool Nonzero(XElement source, string name) {
        string? value = source.Element(source.Name.Namespace + name)?.Value;
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (decimal.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number)) return number != 0;
        try { return System.Xml.XmlConvert.ToTimeSpan(value) != TimeSpan.Zero; }
        catch (FormatException) { return true; }
        catch (OverflowException) { return true; }
    }
}
