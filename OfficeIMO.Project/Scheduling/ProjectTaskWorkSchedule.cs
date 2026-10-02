namespace OfficeIMO.Project;

/// <summary>Calculated task totals. Duration, work and physical completion remain independent measures.</summary>
public sealed class ProjectTaskWorkSchedule {
    internal ProjectTaskWorkSchedule(ProjectWork work, ProjectWork actualWork, ProjectWork remainingWork,
        decimal actualDuration, decimal remainingDuration, decimal? cost, decimal? actualCost, int? physical, bool elapsed = false, bool completed = false, bool hasActuals = false) {
        HasActuals = hasActuals;
        Work = work; ActualWork = actualWork; RemainingWork = remainingWork;
        ActualDuration = ProjectDuration.FromMinutes(actualDuration, ProjectDurationUnit.Minute, elapsed, false, 1);
        RemainingDuration = ProjectDuration.FromMinutes(remainingDuration, ProjectDurationUnit.Minute, elapsed, false, 1);
        Cost = cost; ActualCost = actualCost; RemainingCost = cost - actualCost;
        var actualTime = ProjectWork.FromMinutes(actualDuration);
        var totalTime = ProjectWork.Add(actualTime, ProjectWork.FromMinutes(remainingDuration));
        PercentComplete = totalTime.Minutes == 0 && completed ? 100 : ProjectTimeUnits.Percentage(actualTime, totalTime);
        PercentWorkComplete = ProjectTimeUnits.Percentage(actualWork, work); PhysicalPercentComplete = physical;
    }
    internal bool HasActuals { get; }
    /// <summary>Work-resource assignment total, excluding material quantities.</summary>
    public ProjectWork Work { get; }
    /// <summary>Completed work-resource effort.</summary>
    public ProjectWork ActualWork { get; }
    /// <summary>Remaining work-resource effort.</summary>
    public ProjectWork RemainingWork { get; }
    /// <summary>Completed working duration, separate from effort.</summary>
    public ProjectDuration ActualDuration { get; }
    /// <summary>Remaining working duration, excluding splits and non-working gaps.</summary>
    public ProjectDuration RemainingDuration { get; }
    /// <summary>Duration-based completion percentage.</summary>
    public int PercentComplete { get; }
    /// <summary>Effort-based completion percentage.</summary>
    public int PercentWorkComplete { get; }
    /// <summary>Caller-supplied physical completion; never inferred from dates or effort.</summary>
    public int? PhysicalPercentComplete { get; }
    /// <summary>Calculated assignment or child costs plus fixed cost, adjusted to retain entered task actual cost unless recalculation is enabled.</summary>
    public decimal? Cost { get; }
    /// <summary>Retained task actual cost, or calculated assignment or child actual costs plus accrued fixed cost.</summary>
    public decimal? ActualCost { get; }
    /// <summary>Remaining calculated cost.</summary>
    public decimal? RemainingCost { get; }
}
