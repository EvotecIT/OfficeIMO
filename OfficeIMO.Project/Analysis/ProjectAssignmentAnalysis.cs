using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Stored assignment aggregates for one resource, kept separate from cached resource totals.</summary>
public sealed class ProjectResourceTotals {
    internal ProjectResourceTotals(ProjectResource resource, ProjectAssignment[] assignments) {
        ResourceUid = resource.Uid; AssignmentCount = assignments.Length;
        Cost = Sum(assignments.Select(a => a.Cost)); ActualCost = Sum(assignments.Select(a => a.ActualCost)); RemainingCost = Sum(assignments.Select(a => a.RemainingCost));
        WorkMinutes = Sum(assignments.Select(a => a.Work?.Minutes)); ActualWorkMinutes = Sum(assignments.Select(a => a.ActualWork?.Minutes));
        RemainingWorkMinutes = Sum(assignments.Select(a => a.RemainingWork?.Minutes));
        StoredResourceCost = resource.Cost; StoredResourceActualCost = resource.ActualCost;
    }
    private static decimal? Sum(IEnumerable<decimal?> values) {
        decimal sum = 0; foreach (var value in values) { if (!value.HasValue) return null; sum = checked(sum + value.Value); } return sum;
    }
    /// <summary>Resource identity.</summary>
    public int ResourceUid { get; }
    /// <summary>Number of modeled assignments contributing to these totals.</summary>
    public int AssignmentCount { get; }
    /// <summary>Sum of stored assignment costs; null if any contributor is unspecified.</summary>
    public decimal? Cost { get; }
    /// <summary>Sum of stored actual costs; null if any contributor is unspecified.</summary>
    public decimal? ActualCost { get; }
    /// <summary>Sum of stored remaining costs; null if any contributor is unspecified.</summary>
    public decimal? RemainingCost { get; }
    /// <summary>Sum of stored work minutes; null if any contributor is unspecified.</summary>
    public decimal? WorkMinutes { get; }
    /// <summary>Sum of stored actual work minutes; null if any contributor is unspecified.</summary>
    public decimal? ActualWorkMinutes { get; }
    /// <summary>Sum of stored remaining work minutes; null if any contributor is unspecified.</summary>
    public decimal? RemainingWorkMinutes { get; }
    /// <summary>Independently stored/cached resource cost from the source or model.</summary>
    public decimal? StoredResourceCost { get; }
    /// <summary>Independently stored/cached resource actual cost.</summary>
    public decimal? StoredResourceActualCost { get; }
}

/// <summary>A uniform-rate estimate and actual/remaining consistency assessment for one assignment.</summary>
public sealed class ProjectAssignmentEstimate {
    internal ProjectAssignmentEstimate(ProjectAssignment assignment, decimal? work, decimal? cost) {
        AssignmentUid = assignment.Uid; EstimatedWorkMinutes = work; EstimatedCost = cost;
        StoredCost = assignment.Cost; StoredWorkMinutes = assignment.Work?.Minutes;
    }
    /// <summary>Assignment identity.</summary>
    public int AssignmentUid { get; }
    /// <summary>Stored work, or duration-times-units when a uniform working-duration estimate is possible.</summary>
    public decimal? EstimatedWorkMinutes { get; }
    /// <summary>Uniform-rate estimate. Cost resources retain the explicitly entered amount; unsupported/missing rules yield null.</summary>
    public decimal? EstimatedCost { get; }
    /// <summary>Independently stored cost, never overwritten by analysis.</summary>
    public decimal? StoredCost { get; }
    /// <summary>Independently stored work, never overwritten by analysis.</summary>
    public decimal? StoredWorkMinutes { get; }
}

/// <summary>Immutable resource and assignment analysis bound to one document revision.</summary>
public sealed class ProjectAssignmentAnalysis {
    internal ProjectAssignmentAnalysis(long revision, IEnumerable<ProjectResourceTotals> resources, IEnumerable<ProjectAssignmentEstimate> assignments, IEnumerable<ProjectDiagnostic> diagnostics) {
        ModelRevision = revision; Resources = new ReadOnlyCollection<ProjectResourceTotals>(resources.ToArray());
        Assignments = new ReadOnlyCollection<ProjectAssignmentEstimate>(assignments.ToArray()); Report = new ProjectReport(revision, diagnostics);
    }
    /// <summary>Revision at analysis time.</summary>
    public long ModelRevision { get; }
    /// <summary>Stored assignment aggregates, including discrepancies with source caches.</summary>
    public IReadOnlyList<ProjectResourceTotals> Resources { get; }
    /// <summary>Per-assignment estimates and retained amounts.</summary>
    public IReadOnlyList<ProjectAssignmentEstimate> Assignments { get; }
    /// <summary>Inconsistent actual/remaining values, stale caches, and unsupported estimate rules.</summary>
    public ProjectReport Report { get; }
}
