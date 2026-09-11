using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Explicitly maps one project's work resource to a pool resource. No names, GUIDs, or files are resolved implicitly.</summary>
public sealed class ProjectResourcePoolBinding {
    /// <summary>Creates a mapping using an already calculated assignment schedule. Documents remain caller-owned.</summary>
    public ProjectResourcePoolBinding(ProjectDocument project, ProjectScheduleResult schedule, int resourceUid, int poolResourceUid) {
        Project = project ?? throw new ArgumentNullException(nameof(project));
        Schedule = schedule ?? throw new ArgumentNullException(nameof(schedule));
        ResourceUid = resourceUid; PoolResourceUid = poolResourceUid;
    }
    /// <summary>Project contributing assignment demand.</summary>
    public ProjectDocument Project { get; }
    /// <summary>Immutable assignment schedule for the contributing project's current revision.</summary>
    public ProjectScheduleResult Schedule { get; }
    /// <summary>Resource identity in the contributing project.</summary>
    public int ResourceUid { get; }
    /// <summary>Resource identity in the pool document whose dated availability is authoritative.</summary>
    public int PoolResourceUid { get; }
}

/// <summary>Project-qualified identity of an assignment contributing to shared resource demand.</summary>
public sealed class ProjectResourcePoolAssignment {
    internal ProjectResourcePoolAssignment(ProjectResourcePoolBinding binding, int uid) { Binding = binding; AssignmentUid = uid; }
    /// <summary>Explicit resource mapping and originating project.</summary>
    public ProjectResourcePoolBinding Binding { get; }
    /// <summary>Stable assignment identity within Binding.Project; identities may repeat across projects.</summary>
    public int AssignmentUid { get; }
}

/// <summary>Combined demand from the explicitly bound projects for one pool resource interval.</summary>
public sealed class ProjectResourcePoolInterval {
    internal ProjectResourcePoolInterval(ProjectResourceAllocationInterval interval, IReadOnlyList<ProjectResourcePoolAssignment> assignments) {
        PoolResourceUid = interval.ResourceUid; Start = interval.Start; Finish = interval.Finish;
        Units = interval.Units; Capacity = interval.Capacity;
        Assignments = new ReadOnlyCollection<ProjectResourcePoolAssignment>(interval.AssignmentUids.Select(i => assignments[i]).ToArray());
    }
    /// <summary>Resource identity within the pool.</summary>
    public int PoolResourceUid { get; }
    /// <summary>Inclusive local interval start.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive local interval finish.</summary>
    public DateTime Finish { get; }
    /// <summary>Combined regular assignment demand.</summary>
    public decimal Units { get; }
    /// <summary>Pool MaxUnits or dated availability, without rewriting contributing projects' rates or calendars.</summary>
    public decimal Capacity { get; }
    /// <summary>Positive demand beyond the pool capacity.</summary>
    public decimal ExcessUnits => Math.Max(0m, Units - Capacity);
    /// <summary>Project-qualified contributing assignment identities.</summary>
    public IReadOnlyList<ProjectResourcePoolAssignment> Assignments { get; }
}

/// <summary>Immutable capacity projection of explicitly bound schedules; does not reschedule projects or persist native pool links.</summary>
public sealed class ProjectResourcePoolResult {
    internal ProjectResourcePoolResult(long revision, IEnumerable<ProjectResourcePoolInterval> intervals) {
        ModelRevision = revision; Intervals = new ReadOnlyCollection<ProjectResourcePoolInterval>(intervals.ToArray());
        Overallocations = new ReadOnlyCollection<ProjectResourcePoolInterval>(Intervals.Where(i => i.ExcessUnits > .000000001m).ToArray());
    }
    /// <summary>Pool revision supplying capacity.</summary>
    public long ModelRevision { get; }
    /// <summary>Occupied intervals for the supplied bindings only. Unbound resources do not contribute.</summary>
    public IReadOnlyList<ProjectResourcePoolInterval> Intervals { get; }
    /// <summary>Intervals whose combined demand exceeds pool availability.</summary>
    public IReadOnlyList<ProjectResourcePoolInterval> Overallocations { get; }
}
