namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Analyzes concurrent regular work against dated resource availability, without moving tasks or treating overtime as regular capacity.</summary>
    public ProjectResourceAllocationResult AnalyzeResourceAllocation(ProjectScheduleResult schedule, int maxIntervals = 1_000_000, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (schedule == null) throw new ArgumentNullException(nameof(schedule));
        if (maxIntervals < 1) throw new ArgumentOutOfRangeException(nameof(maxIntervals));
        if (_batchDepth != 0 || schedule.Document != this || schedule.ModelRevision != Revision)
            throw new InvalidOperationException("Capacity analysis requires a schedule for the current document revision.");
        if (!schedule.CalculatedAssignments) throw new ArgumentException("Capacity analysis requires CalculateAssignments.", nameof(schedule));
        schedule.Report.ThrowIfErrors(); cancellationToken.ThrowIfCancellationRequested();
        foreach (var source in schedule.ExternalSources) source.ValidateCurrent();
        var result = ProjectResourceAllocation.Analyze(this, schedule, maxIntervals, cancellationToken);
        if (schedule.ModelRevision != Revision) throw new InvalidOperationException("The document changed during capacity analysis.");
        foreach (var source in schedule.ExternalSources) source.ValidateCurrent();
        return result;
    }
}
