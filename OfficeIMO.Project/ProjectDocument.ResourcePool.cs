namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>
    /// Combines explicitly mapped assignment demand against this pool's dated availability, without file access or mutation.
    /// Schedules use the same local time basis. Rates, calendars, and resource names are not synchronized; calculate each source first.
    /// Include the pool's own assignments through an explicit binding when they should contribute.
    /// </summary>
    public ProjectResourcePoolResult AnalyzeResourcePool(IEnumerable<ProjectResourcePoolBinding> bindings,
        int maxIntervals = 1_000_000, int maxBindings = 10_000, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (bindings == null) throw new ArgumentNullException(nameof(bindings));
        if (maxIntervals < 1) throw new ArgumentOutOfRangeException(nameof(maxIntervals));
        if (maxBindings < 1) throw new ArgumentOutOfRangeException(nameof(maxBindings));
        if (HasActiveUpdate) throw new InvalidOperationException("Pool analysis requires completed document updates.");
        long revision = Revision;
        var selected = new List<ProjectResourcePoolBinding>();
        var identities = new HashSet<(ProjectDocument, int)>();
        foreach (var binding in bindings) {
            cancellationToken.ThrowIfCancellationRequested();
            if (selected.Count >= maxBindings) throw new InvalidOperationException("Resource pool analysis exceeds MaxBindings.");
            if (binding == null) throw new ArgumentException("A resource pool binding is null.", nameof(bindings));
            ValidatePoolSchedule(binding);
            if (!identities.Add((binding.Project, binding.ResourceUid)))
                throw new ArgumentException("A contributing resource may be bound only once.", nameof(bindings));
            if (binding.Project.Resources.GetByUid(binding.ResourceUid).Type != ProjectResourceType.Work || Resources.GetByUid(binding.PoolResourceUid).Type != ProjectResourceType.Work)
                throw new NotSupportedException("Resource pools analyze work resources only.");
            selected.Add(binding);
        }
        Validate(cancellationToken).ThrowIfErrors();
        var assignments = new List<ProjectResourcePoolAssignment>();
        var intervals = new List<ProjectResourceAllocationInterval>();
        long inputIntervals = 0;
        foreach (var group in selected.GroupBy(b => b.PoolResourceUid).OrderBy(g => g.Key)) {
            var demand = new List<(int AssignmentUid, IReadOnlyList<ProjectAssignmentInterval> Intervals)>();
            foreach (var binding in group) foreach (var assignment in binding.Schedule.Assignments.Where(a => a.ResourceUid == binding.ResourceUid)) {
                cancellationToken.ThrowIfCancellationRequested();
                inputIntervals += assignment.Intervals.Count;
                if (inputIntervals > maxIntervals || assignments.Count >= maxIntervals)
                    throw new InvalidOperationException("Resource pool analysis exceeds MaxIntervals.");
                demand.Add((assignments.Count, assignment.Intervals));
                assignments.Add(new ProjectResourcePoolAssignment(binding, assignment.AssignmentUid));
            }
            ProjectResourceAllocation.AnalyzeResource(Resources.GetByUid(group.Key), demand, intervals, maxIntervals, cancellationToken);
        }
        foreach (var binding in selected) ValidatePoolSchedule(binding);
        EnsureNotDisposed(); cancellationToken.ThrowIfCancellationRequested();
        if (Revision != revision || HasActiveUpdate) throw new InvalidOperationException("The pool changed during capacity analysis.");
        return new ProjectResourcePoolResult(revision, intervals.Select(i => new ProjectResourcePoolInterval(i, assignments)));
    }

    private static void ValidatePoolSchedule(ProjectResourcePoolBinding binding) {
        binding.Project.EnsureNotDisposed();
        if (binding.Project.HasActiveUpdate || binding.Schedule.Document != binding.Project || binding.Schedule.ModelRevision != binding.Project.Revision)
            throw new InvalidOperationException("Pool analysis requires current schedules and completed document updates.");
        if (!binding.Schedule.CalculatedAssignments) throw new ArgumentException("Pool analysis requires CalculateAssignments.");
        binding.Schedule.Report.ThrowIfErrors();
        foreach (var source in binding.Schedule.ExternalSources) source.ValidateCurrent();
    }
}
