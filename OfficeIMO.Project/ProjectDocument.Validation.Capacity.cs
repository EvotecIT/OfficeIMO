namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private void CheckResourceCapacity(Finding add, CancellationToken token) {
        foreach (var task in AllTasks) {
            token.ThrowIfCancellationRequested(); string path = "/Task[UID=" + task.Uid + "]";
            CheckEnum(task.FixedCostAccrual, path + "/FixedCostAccrual", add); CheckEnum(task.EarnedValueMethod, path + "/EarnedValueMethod", add);
            CheckDateRange(task.Stop, task.Resume, path + "/Progress", add);
            if (task.LevelingDelay is ProjectDuration delay && !IsValidLevelingDelay(delay))
                add("PROJECT_LEVELING_DELAY", "Leveling delay must be nonnegative whole tenths of a minute.", path);
        }
        foreach (var resource in Resources) {
            token.ThrowIfCancellationRequested(); string path = "/Resource[UID=" + resource.Uid + "]";
            CheckEnum(resource.AccrueAt, path + "/AccrueAt", add);
            if (resource.StandardRate < 0 || resource.OvertimeRate < 0 || resource.CostPerUse < 0)
                add("PROJECT_RATE_VALUE", "Rates and per-use charges cannot be negative.", path);
            DateTime? previousEnd = null; bool previous = false;
            foreach (var period in resource.AvailabilityPeriods.OrderBy(p => p.From ?? DateTime.MinValue)) {
                token.ThrowIfCancellationRequested();
                CheckDateRange(period.From, period.Through, path + "/Availability", add);
                if ((period.From?.Ticks ?? 0) % TimeSpan.TicksPerMinute != 0 || (period.Through?.Ticks ?? 0) % TimeSpan.TicksPerMinute != 0)
                    add("PROJECT_AVAILABILITY_PRECISION", "Availability bounds use whole local minutes.", path + "/Availability");
                if (previous && (!previousEnd.HasValue || !period.From.HasValue || period.From <= previousEnd))
                    add("PROJECT_AVAILABILITY_OVERLAP", "Resource availability periods cannot overlap.", path + "/Availability");
                previousEnd = period.Through; previous = true;
            }
            foreach (var table in resource.Rates.GroupBy(r => r.Table ?? ProjectCostRateTable.A)) {
                previousEnd = null;
                foreach (var rate in table.OrderBy(r => r.From)) {
                    token.ThrowIfCancellationRequested();
                    CheckEnum(rate.Table, path + "/Rate/Table", add); CheckDateRange(rate.From, rate.To, path + "/Rate", add);
                    if (!rate.From.HasValue || !rate.To.HasValue || rate.From >= rate.To)
                        add("PROJECT_RATE_PERIOD", "Dated rates require an explicit nonempty half-open range.", path + "/Rate");
                    if (previousEnd > rate.From) add("PROJECT_RATE_OVERLAP", "Rates in one table cannot overlap.", path + "/Rate");
                    if (rate.StandardRate < 0 || rate.OvertimeRate < 0 || rate.CostPerUse < 0)
                        add("PROJECT_RATE_VALUE", "Rates and per-use charges cannot be negative.", path + "/Rate");
                    previousEnd = rate.To;
                }
            }
        }
        foreach (var assignment in Assignments) {
            token.ThrowIfCancellationRequested(); string path = "/Assignment[UID=" + assignment.Uid + "]";
            CheckEnum(assignment.CostRateTable, path + "/CostRateTable", add); CheckEnum(assignment.WorkContour, path + "/WorkContour", add);
            CheckDateRange(assignment.Stop, assignment.Resume, path + "/Progress", add);
            if (assignment.DelayMinutes is decimal assignmentDelay && !IsValidTenths(assignmentDelay))
                add("PROJECT_ASSIGNMENT_DELAY", "Assignment delay must be nonnegative whole tenths of a working minute.", path);
            if (assignment.HasFixedRateUnits == false && assignment.Resource?.Type == ProjectResourceType.Material &&
                (assignment.MaterialRateScale < 1 || assignment.MaterialRateScale > 5 || !assignment.MaterialRateScale.HasValue))
                add("PROJECT_MATERIAL_RATE_SCALE", "Variable material usage requires a minute, hour, day, week or month rate scale (1–5).", path);
        }
    }
    private bool IsValidLevelingDelay(ProjectDuration delay) {
        if (delay.Value < 0) return false;
        try {
            decimal minutes = checked(delay.Value * ProjectXmlValue.MinutesPerUnit(delay.Unit, delay.IsElapsed, this));
            return ProjectXmlValue.CanRepresentMinutes(minutes) && checked(minutes * 10m) % 1m == 0;
        } catch (OverflowException) { return false; }
    }
    private static bool IsValidTenths(decimal value) {
        if (value < 0) return false;
        try { return checked(value * 10m) % 1m == 0; }
        catch (OverflowException) { return false; }
    }
}
