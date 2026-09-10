using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Deterministic bounds for an explicit schedule calculation.</summary>
public sealed class ProjectScheduleOptions {
    /// <summary>Maximum calendar dates visited by an arithmetic operation; prevents infinite searches in closed calendars.</summary>
    public int MaxCalendarDays { get; set; } = 36600;
    /// <summary>Total float at or below this many working minutes is critical.</summary>
    public decimal CriticalSlackMinutes { get; set; }
}

/// <summary>Calculated dates and float for one task. Values never mutate the stored task during calculation.</summary>
public sealed class ProjectTaskSchedule {
    internal ProjectTaskSchedule(int uid, DateTime start, DateTime finish, DateTime earlyStart, DateTime earlyFinish,
        DateTime lateStart, DateTime lateFinish, decimal total, decimal free, bool critical, bool summary, ProjectDuration duration) {
        TaskUid = uid; Start = start; Finish = finish; EarlyStart = earlyStart; EarlyFinish = earlyFinish;
        LateStart = lateStart; LateFinish = lateFinish; TotalSlackMinutes = total; FreeSlackMinutes = free; IsCritical = critical; IsSummary = summary;
        Duration = duration;
    }
    /// <summary>Stable task identity in the originating document.</summary>
    public int TaskUid { get; }
    /// <summary>Proposed start for the selected forward/backward policy.</summary>
    public DateTime Start { get; }
    /// <summary>Proposed finish.</summary>
    public DateTime Finish { get; }
    /// <summary>Calculated duration, including the resolved fixed-work equation or summary span.</summary>
    public ProjectDuration Duration { get; }
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
    internal ProjectScheduleResult(ProjectDocument document, long revision, IEnumerable<ProjectTaskSchedule> tasks, IEnumerable<ProjectDiagnostic> diagnostics) {
        Document = document; ModelRevision = revision;
        Tasks = new ReadOnlyCollection<ProjectTaskSchedule>(tasks.ToArray());
        Report = new ProjectReport(revision, diagnostics);
    }
    /// <summary>Revision used for calculation; applying to a different revision is rejected.</summary>
    public long ModelRevision { get; }
    /// <summary>Proposed task results, in outline order. Empty when input rules prevent calculation.</summary>
    public IReadOnlyList<ProjectTaskSchedule> Tasks { get; }
    /// <summary>Structural errors, unsupported combinations, and constraint conflicts.</summary>
    public ProjectReport Report { get; }
}
