namespace OfficeIMO.Project;

/// <summary>Explicit, calendar-independent work/units/duration equations for uniform work-resource assignments.</summary>
public static class ProjectWorkEquation {
    /// <summary>Computes work from working minutes and allocation units. Elapsed durations must first be resolved against a calendar.</summary>
    public static ProjectWork Work(decimal workingMinutes, ProjectUnits units) {
        if (workingMinutes < 0) throw new ArgumentOutOfRangeException(nameof(workingMinutes));
        return new ProjectWork(checked(workingMinutes * units.Value));
    }
    /// <summary>Computes the uniform working duration for a fixed amount of work and positive total units.</summary>
    public static decimal DurationMinutes(ProjectWork work, ProjectUnits totalUnits) {
        if (totalUnits.Value <= 0) throw new ArgumentOutOfRangeException(nameof(totalUnits), "Fixed work requires positive allocation units.");
        return work.Minutes / totalUnits.Value;
    }
    /// <summary>Computes allocation for fixed work and duration. Zero-duration work cannot be inferred.</summary>
    public static ProjectUnits Units(ProjectWork work, decimal workingMinutes) {
        if (workingMinutes <= 0) throw new ArgumentOutOfRangeException(nameof(workingMinutes));
        return ProjectUnits.Fraction(work.Minutes / workingMinutes);
    }
    /// <summary>Uniform planned work cost: regular hours times standard hourly rate plus overtime hours times overtime hourly rate plus the assignment's per-use charge, already scaled by work-resource units.</summary>
    public static decimal WorkCost(ProjectWork totalWork, ProjectWork overtimeWork, decimal standardHourlyRate, decimal overtimeHourlyRate, decimal costPerUse = 0) {
        if (overtimeWork.Minutes > totalWork.Minutes) throw new ArgumentException("Overtime is part of total work, not additional to it.", nameof(overtimeWork));
        if (standardHourlyRate < 0 || overtimeHourlyRate < 0 || costPerUse < 0) throw new ArgumentOutOfRangeException(nameof(standardHourlyRate));
        return checked((totalWork.Minutes - overtimeWork.Minutes) / 60m * standardHourlyRate + overtimeWork.Minutes / 60m * overtimeHourlyRate + costPerUse);
    }
    /// <summary>Material quantity times price per material unit plus one per-use charge; no work-resource percentage conversion.</summary>
    public static decimal MaterialCost(decimal quantity, decimal unitPrice, decimal costPerUse = 0) {
        if (quantity < 0 || unitPrice < 0 || costPerUse < 0) throw new ArgumentOutOfRangeException(nameof(quantity));
        return checked(quantity * unitPrice + costPerUse);
    }
}
