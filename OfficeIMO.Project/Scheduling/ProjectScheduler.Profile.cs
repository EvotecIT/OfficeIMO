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
            if (assignment.TimephasedData.Any(v => v.Type == 1 && (!v.Start.HasValue || !v.Finish.HasValue || v.Start < result.Start || v.Finish > result.Finish)))
                Error("PROJECT_ASSIGNMENT_DATE_RECALCULATION_REQUIRED", "Stored remaining-work curves do not fit the proposed task dates. Calculate assignments before applying these dates.", assignment.Task);
        }
    }
    private void CheckSourceProfile() {
        if (_document.MpxSource?.Unmodeled.Count > 0)
            Error("PROJECT_MPX_SCHEDULING_PROFILE", "MPX contains unmodeled fields or records that may affect scheduling. Supply a fully typed scheduling model before calculation.");
        if (_document.AllTasks.Any(task => task.IsRecurring == true))
            _diagnostics.Add(new ProjectDiagnostic("PROJECT_EXPANDED_RECURRENCE", ProjectDiagnosticSeverity.Warning,
                "Calculation schedules the stored recurring occurrences independently. It does not infer or regenerate a recurrence rule from the series marker.", "/Project"));
        if (_options.CalculateAssignments && _document.NativeSource != null)
            Error("PROJECT_NATIVE_ASSIGNMENT_PROFILE", "Independent assignment calculation requires typed timephased and rate inputs; opaque native assignment profiles cannot be calculated.");
        if (_document.NativeSource != null)
            _diagnostics.Add(new ProjectDiagnostic("PROJECT_NATIVE_SCHEDULE_PROJECTION", ProjectDiagnosticSeverity.Warning,
                "Calculation uses decoded scalar values and calendars. Native contours, splits, leveling, and other opaque scheduling records are not interpreted; this is a projection of the supported model.", "/Project"));
        foreach (var task in _document.AllTasks) {
            _token.ThrowIfCancellationRequested();
            if (!_options.CalculateAssignments && task.IgnoreResourceCalendar == true)
                Error("PROJECT_TASK_SCHEDULING_PROFILE", "Ignoring resource calendars requires independent assignment calculation.", task);
            if (task.LevelingDelay?.Value > 0 && _document.Settings.ScheduleFromStart == false)
                Error("PROJECT_LEVELING_BACKWARD", "Preserved leveling delays require forward scheduling; clear them explicitly before scheduling backward.", task);
            var source = _document.Source?.Element(task);
            if (source == null) continue;
            if (Enabled(source, "ExternalTask") || Enabled(source, "IsSubproject") ||
                (!_options.CalculateAssignments && Enabled(source, "IgnoreResourceCalendar")) || source.Element(source.Name.Namespace + "RecurringTask") != null)
                Error("PROJECT_TASK_SCHEDULING_PROFILE", "External tasks, opaque recurrence rules, and ignored resource calendars require additional scheduling semantics.", task);
        }
        foreach (var assignment in _document.Assignments) {
            _token.ThrowIfCancellationRequested();
            if (_options.CalculateAssignments && assignment.Task?.IsSummary == true)
                Error("PROJECT_SUMMARY_ASSIGNMENT_PROFILE", "Direct summary-task assignments are not part of independent assignment calculation. Assign resources to leaf tasks before calculating summary rollups.", assignment.Task);
            if (!_options.CalculateAssignments && (assignment.DelayMinutes > 0 || assignment.WorkContour.HasValue && assignment.WorkContour != ProjectWorkContour.Flat))
                Error("PROJECT_ASSIGNMENT_SCHEDULING_PROFILE", "Assignment delays and non-flat contours require independent assignment scheduling.", assignment.Task);
            if (!_options.CalculateAssignments && (assignment.ActualStart.HasValue || assignment.ActualFinish.HasValue || assignment.ActualWork?.Minutes > 0 || assignment.PercentWorkComplete > 0
                || assignment.TimephasedData.Any(v => v.Type == 2 || v.Type == 3)))
                Error("PROJECT_PROGRESS_SCHEDULING", "Assignment progress requires interval-aware rescheduling.", assignment.Task);
            var source = _document.Source?.Element(assignment);
            if (source != null && ((!_options.CalculateAssignments && (Nonzero(source, "Delay") || Nonzero(source, "WorkContour"))) || Nonzero(source, "LevelingDelay")))
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
