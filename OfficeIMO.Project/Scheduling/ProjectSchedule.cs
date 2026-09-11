using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Deterministic bounds for an explicit schedule calculation.</summary>
public sealed class ProjectScheduleOptions {
    /// <summary>Calculate independent assignment intervals, work, progress and costs. False retains the date-only profile.</summary>
    public bool CalculateAssignments { get; set; }
    /// <summary>Redistribute an effort-driven task's remaining work across its current assignments, preserving completed work.</summary>
    public bool RedistributeEffortDrivenWork { get; set; }
    /// <summary>Move unfinished work that precedes the explicitly stored StatusDate to that boundary.</summary>
    public bool RescheduleRemainingAfterStatusDate { get; set; }
    /// <summary>Recompute actual costs from recorded actual work and rates. Otherwise stored actual costs are retained.</summary>
    public bool RecalculateActualCosts { get; set; }
    /// <summary>Maximum materialized assignment and cost intervals for one calculation.</summary>
    public int MaxIntervals { get; set; } = 1_000_000;
    /// <summary>Maximum calendar dates visited by an arithmetic operation; prevents infinite searches in closed calendars.</summary>
    public int MaxCalendarDays { get; set; } = 36600;
    /// <summary>Optional caller-controlled resolver for external Project XML predecessor references. The engine never opens external files itself.</summary>
    public ProjectExternalProjectResolver? ExternalProjectResolver { get; set; }
    /// <summary>Maximum distinct referring-project/reference pairs resolved during one calculation.</summary>
    public int MaxExternalProjects { get; set; } = 32;
    /// <summary>Maximum external dependency depth below the local project.</summary>
    public int MaxExternalDepth { get; set; } = 16;
    /// <summary>Maximum tasks retained across resolved external schedule proposals.</summary>
    public int MaxExternalTasks { get; set; } = 100000;
    /// <summary>Total float at or below this many working minutes is critical.</summary>
    public decimal CriticalSlackMinutes { get; set; }
}

/// <summary>Calculated dates and float for one task. Values never mutate the stored task during calculation.</summary>
public sealed class ProjectTaskSchedule {
    internal ProjectTaskSchedule(int uid, DateTime start, DateTime finish, DateTime earlyStart, DateTime earlyFinish,
        DateTime lateStart, DateTime lateFinish, decimal total, decimal free, bool critical, bool summary, ProjectDuration duration, ProjectTaskWorkSchedule? calculation = null,
        DateTime beforeLevelingAnchor = default, DateTime forwardAnchor = default) {
        TaskUid = uid; Start = start; Finish = finish; EarlyStart = earlyStart; EarlyFinish = earlyFinish;
        LateStart = lateStart; LateFinish = lateFinish; TotalSlackMinutes = total; FreeSlackMinutes = free; IsCritical = critical; IsSummary = summary;
        Duration = duration;
        Calculation = calculation;
        BeforeLevelingAnchor = beforeLevelingAnchor; ForwardAnchor = forwardAnchor;
    }
    // Keep the dependency/constraint bound distinct from actual start and assignment delays.
    internal DateTime BeforeLevelingAnchor { get; }
    internal DateTime ForwardAnchor { get; }
    /// <summary>Stable task identity in the originating document.</summary>
    public int TaskUid { get; }
    /// <summary>Proposed start for the selected forward/backward policy.</summary>
    public DateTime Start { get; }
    /// <summary>Proposed finish.</summary>
    public DateTime Finish { get; }
    /// <summary>Calculated duration, including the resolved fixed-work equation or summary span.</summary>
    public ProjectDuration Duration { get; }
    /// <summary>Work, progress and cost totals when independent assignment calculation was requested.</summary>
    public ProjectTaskWorkSchedule? Calculation { get; }
    /// <summary>Earliest feasible start from the forward pass.</summary>
    public DateTime EarlyStart { get; }
    /// <summary>Earliest finish.</summary>
    public DateTime EarlyFinish { get; }
    /// <summary>Latest start respecting successors and the project finish/deadline bounds.</summary>
    public DateTime LateStart { get; }
    /// <summary>Latest finish.</summary>
    public DateTime LateFinish { get; }
    /// <summary>Signed working minutes between early and late start. Negative values expose an infeasible target.</summary>
    public decimal TotalSlackMinutes { get; }
    /// <summary>Working minutes this task can move without moving a successor's current bound.</summary>
    public decimal FreeSlackMinutes { get; }
    /// <summary>Whether total float meets the selected critical threshold.</summary>
    public bool IsCritical { get; }
    /// <summary>Whether this entry rolls up descendant dates.</summary>
    public bool IsSummary { get; }
}

/// <summary>Immutable schedule proposal bound to one document and mutation revision.</summary>
public sealed class ProjectScheduleResult {
    internal readonly ProjectDocument Document;
    internal ProjectScheduleResult(ProjectDocument document, long revision, IEnumerable<ProjectTaskSchedule> tasks, IEnumerable<ProjectDiagnostic> diagnostics,
        IEnumerable<ProjectAssignmentSchedule>? assignments = null, bool calculatedAssignments = false, IEnumerable<ProjectExternalScheduleSource>? externalSources = null) {
        Document = document; ModelRevision = revision;
        Tasks = new ReadOnlyCollection<ProjectTaskSchedule>(tasks.ToArray());
        Report = new ProjectReport(revision, diagnostics);
        Assignments = new ReadOnlyCollection<ProjectAssignmentSchedule>((assignments ?? Array.Empty<ProjectAssignmentSchedule>()).ToArray());
        CalculatedAssignments = calculatedAssignments;
        ExternalSources = new ReadOnlyCollection<ProjectExternalScheduleSource>((externalSources ?? Array.Empty<ProjectExternalScheduleSource>()).ToArray());
    }
    /// <summary>External projects contributing predecessor bounds. The result's task and assignment collections remain local; resource-pool analysis is a separate operation.</summary>
    public IReadOnlyList<ProjectExternalScheduleSource> ExternalSources { get; }
    /// <summary>Revision used for calculation; applying to a different revision is rejected.</summary>
    public long ModelRevision { get; }
    /// <summary>Proposed task results, in outline order. Empty when input rules prevent calculation.</summary>
    public IReadOnlyList<ProjectTaskSchedule> Tasks { get; }
    /// <summary>Independent assignment dates, work and costs; empty for the date-only profile.</summary>
    public IReadOnlyList<ProjectAssignmentSchedule> Assignments { get; }
    /// <summary>Whether calculation used independent assignment and progress semantics.</summary>
    public bool CalculatedAssignments { get; }
    /// <summary>Structural errors, unsupported combinations, and constraint conflicts.</summary>
    public ProjectReport Report { get; }
}
