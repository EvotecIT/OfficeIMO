using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Baseline budget, earned value, and actual costs for one task at an explicit status boundary.</summary>
public sealed class ProjectTaskEarnedValue {
    internal ProjectTaskEarnedValue(int uid, decimal? budget, decimal? planned, decimal? earned, decimal? actual) {
        TaskUid = uid; BudgetAtCompletion = budget; PlannedValue = planned; EarnedValue = earned; ActualCost = actual;
    }
    /// <summary>Task identity.</summary>
    public int TaskUid { get; }
    /// <summary>Selected baseline cost (BAC).</summary>
    public decimal? BudgetAtCompletion { get; }
    /// <summary>Baseline cost scheduled through StatusDate (BCWS or PV).</summary>
    public decimal? PlannedValue { get; }
    /// <summary>Budgeted value of completed work (BCWP or EV).</summary>
    public decimal? EarnedValue { get; }
    /// <summary>Actual costs accrued through StatusDate (ACWP or AC).</summary>
    public decimal? ActualCost { get; }
    /// <summary>Earned value minus actual costs.</summary>
    public decimal? CostVariance => EarnedValue - ActualCost;
    /// <summary>Earned value minus planned value.</summary>
    public decimal? ScheduleVariance => EarnedValue - PlannedValue;
    /// <summary>Earned value divided by actual cost; null when unavailable or zero.</summary>
    public decimal? CostPerformanceIndex => ActualCost.HasValue && ActualCost != 0 ? EarnedValue / ActualCost : null;
    /// <summary>Earned value divided by planned value; null when unavailable or zero.</summary>
    public decimal? SchedulePerformanceIndex => PlannedValue.HasValue && PlannedValue != 0 ? EarnedValue / PlannedValue : null;
    /// <summary>BAC divided by CPI, assuming future work follows observed cost performance.</summary>
    public decimal? EstimateAtCompletion => CostPerformanceIndex.HasValue && CostPerformanceIndex != 0 ? BudgetAtCompletion / CostPerformanceIndex : null;
    /// <summary>Forecast remaining cost under the CPI assumption.</summary>
    public decimal? EstimateToComplete => EstimateAtCompletion - ActualCost;
    /// <summary>Budget minus the CPI-based estimate at completion.</summary>
    public decimal? VarianceAtCompletion => BudgetAtCompletion - EstimateAtCompletion;
}

/// <summary>Immutable earned-value analysis. Missing baseline curves produce diagnostics and unknown values, rather than a fabricated uniform budget.</summary>
public sealed class ProjectEarnedValueResult {
    internal ProjectEarnedValueResult(long revision, DateTime status, int number, IEnumerable<ProjectTaskEarnedValue> tasks, IEnumerable<ProjectDiagnostic> diagnostics) {
        ModelRevision = revision; StatusDate = status; BaselineNumber = number;
        Tasks = new ReadOnlyCollection<ProjectTaskEarnedValue>(tasks.ToArray()); Report = new ProjectReport(revision, diagnostics);
    }
    /// <summary>Analyzed document revision.</summary>
    public long ModelRevision { get; }
    /// <summary>Explicit local cost cutoff; intervals are clipped at this boundary.</summary>
    public DateTime StatusDate { get; }
    /// <summary>Selected baseline, 0 through 10.</summary>
    public int BaselineNumber { get; }
    /// <summary>Task results in outline order.</summary>
    public IReadOnlyList<ProjectTaskEarnedValue> Tasks { get; }
    /// <summary>Missing data, inconsistent costs, or unsupported curve information.</summary>
    public ProjectReport Report { get; }
}

internal static class ProjectBaselineTypes {
    internal static (int Work, int Cost) Task(int number) => number == 0 ? (9, 10) : (18 + (number - 1) * 6, 19 + (number - 1) * 6);
    internal static (int Work, int Cost) Assignment(int number) => number == 0 ? (4, 5) : (16 + (number - 1) * 6, 17 + (number - 1) * 6);
    internal static (int Work, int Cost) Resource(int number) => number == 0 ? (7, 8) : (20 + (number - 1) * 6, 21 + (number - 1) * 6);
}
