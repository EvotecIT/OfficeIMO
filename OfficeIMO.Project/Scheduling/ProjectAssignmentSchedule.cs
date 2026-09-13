using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>One uninterrupted allocation on a resource's effective working calendar.</summary>
public sealed class ProjectAssignmentInterval {
    internal ProjectAssignmentInterval(DateTime start, DateTime finish, decimal work, decimal overtime, bool actual) {
        Start = start; Finish = finish; Work = new ProjectWork(work); OvertimeWork = new ProjectWork(overtime); IsActual = actual;
    }
    /// <summary>Inclusive local start.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive local finish.</summary>
    public DateTime Finish { get; }
    /// <summary>Total work, including overtime. For materials, hours encode material quantity as required by MSPDI.</summary>
    public ProjectWork Work { get; }
    /// <summary>Overtime included in Work.</summary>
    public ProjectWork OvertimeWork { get; }
    /// <summary>True for recorded actual work, false for proposed remaining work.</summary>
    public bool IsActual { get; }
    /// <summary>Regular work divided by interval minutes; material intervals are consumption rates, not resource capacity.</summary>
    public decimal Units => Finish == Start ? 0m : (Work.Minutes - OvertimeWork.Minutes) / ((Finish.Ticks - Start.Ticks) / (decimal)TimeSpan.TicksPerMinute);
}

/// <summary>A currency amount accrued over an interval, or at a single instant when Start equals Finish.</summary>
public sealed class ProjectCostInterval {
    internal ProjectCostInterval(DateTime start, DateTime finish, decimal cost, bool actual) { Start = start; Finish = finish; Cost = cost; IsActual = actual; }
    /// <summary>Inclusive accrual start.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive accrual finish; equal endpoints indicate a one-time charge.</summary>
    public DateTime Finish { get; }
    /// <summary>Amount in the project's currency, without implicit exchange conversion.</summary>
    public decimal Cost { get; }
    /// <summary>Whether the cost is actual rather than remaining.</summary>
    public bool IsActual { get; }
}

/// <summary>Immutable proposed work, material consumption, costs and dates for one assignment.</summary>
public sealed class ProjectAssignmentSchedule {
    internal ProjectAssignmentSchedule(ProjectAssignment assignment, DateTime start, DateTime finish, decimal units,
        IEnumerable<ProjectAssignmentInterval> intervals, IEnumerable<ProjectCostInterval> costs, decimal? cost, decimal? actualCost, decimal? materialQuantity,
        DateTime? remainingAnchor = null) {
        AssignmentUid = assignment.Uid; TaskUid = assignment.Task!.Uid; ResourceUid = assignment.Resource!.Uid;
        Start = start; Finish = finish; Units = ProjectUnits.Fraction(units);
        RemainingAnchor = remainingAnchor ?? start;
        Intervals = new ReadOnlyCollection<ProjectAssignmentInterval>(intervals.ToArray());
        Costs = new ReadOnlyCollection<ProjectCostInterval>(costs.ToArray());
        Work = new ProjectWork(Intervals.Sum(i => i.Work.Minutes)); ActualWork = new ProjectWork(Intervals.Where(i => i.IsActual).Sum(i => i.Work.Minutes));
        RemainingWork = new ProjectWork(Work.Minutes - ActualWork.Minutes);
        OvertimeWork = new ProjectWork(Intervals.Sum(i => i.OvertimeWork.Minutes));
        ActualOvertimeWork = new ProjectWork(Intervals.Where(i => i.IsActual).Sum(i => i.OvertimeWork.Minutes));
        Cost = cost; ActualCost = actualCost; RemainingCost = cost - actualCost; MaterialQuantity = materialQuantity;
    }
    /// <summary>Stable assignment identity in the originating project.</summary>
    public int AssignmentUid { get; }
    /// <summary>Owning task UID.</summary>
    public int TaskUid { get; }
    /// <summary>Assigned resource UID.</summary>
    public int ResourceUid { get; }
    /// <summary>First planned or actual assignment boundary.</summary>
    public DateTime Start { get; }
    /// <summary>Last planned or actual assignment boundary.</summary>
    public DateTime Finish { get; }
    internal DateTime RemainingAnchor { get; }
    /// <summary>Proposed scalar allocation or material consumption setting; interval allocations may vary.</summary>
    public ProjectUnits Units { get; }
    /// <summary>Total work, including overtime.</summary>
    public ProjectWork Work { get; }
    /// <summary>Recorded actual work.</summary>
    public ProjectWork ActualWork { get; }
    /// <summary>Proposed remaining work.</summary>
    public ProjectWork RemainingWork { get; }
    /// <summary>Total overtime included in Work.</summary>
    public ProjectWork OvertimeWork { get; }
    /// <summary>Recorded actual overtime included in ActualWork.</summary>
    public ProjectWork ActualOvertimeWork { get; }
    /// <summary>Calculated total cost; null when the selected rates are incomplete.</summary>
    public decimal? Cost { get; }
    /// <summary>Actual cost, preserving the stored amount unless recalculation was explicitly requested.</summary>
    public decimal? ActualCost { get; }
    /// <summary>Cost minus actual cost, when both amounts are available.</summary>
    public decimal? RemainingCost { get; }
    /// <summary>Total material units consumed, or null for work/cost resources.</summary>
    public decimal? MaterialQuantity { get; }
    /// <summary>Actual and proposed remaining work, split at non-working periods and explicit gaps.</summary>
    public IReadOnlyList<ProjectAssignmentInterval> Intervals { get; }
    /// <summary>Interval and one-time costs used by cost accrual and earned-value analysis.</summary>
    public IReadOnlyList<ProjectCostInterval> Costs { get; }
}
