namespace OfficeIMO.Project;

/// <summary>Calculated task totals. Duration, work and physical completion remain independent measures.</summary>
public sealed class ProjectTaskWorkSchedule {
    internal ProjectTaskWorkSchedule(ProjectWork work, ProjectWork actualWork, ProjectWork remainingWork,
        decimal actualDuration, decimal remainingDuration, decimal? cost, decimal? actualCost, int? physical, bool elapsed = false) {
        Work = work; ActualWork = actualWork; RemainingWork = remainingWork;
        ActualDuration = new ProjectDuration(actualDuration, ProjectDurationUnit.Minute, elapsed); RemainingDuration = new ProjectDuration(remainingDuration, ProjectDurationUnit.Minute, elapsed);
        Cost = cost; ActualCost = actualCost; RemainingCost = cost - actualCost;
        PercentComplete = Percentage(actualDuration, actualDuration + remainingDuration);
        PercentWorkComplete = Percentage(actualWork.Minutes, work.Minutes); PhysicalPercentComplete = physical;
    }
    private static int Percentage(decimal actual, decimal total) => total <= 0 ? 0 : (int)decimal.Round(Math.Min(100, actual / total * 100), 0, MidpointRounding.AwayFromZero);
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
    /// <summary>Total assignment costs plus the task's fixed cost, when rates are complete.</summary>
    public decimal? Cost { get; }
    /// <summary>Actual assignment costs plus accrued fixed cost.</summary>
    public decimal? ActualCost { get; }
    /// <summary>Remaining calculated cost.</summary>
    public decimal? RemainingCost { get; }
}
