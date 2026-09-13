using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>Resolves a declared external project reference. The caller controls all file/network access and retains ownership of returned documents.</summary>
public delegate ProjectDocument? ProjectExternalProjectResolver(ProjectDocument referringProject, string projectReference, CancellationToken cancellationToken);

/// <summary>A project whose calculated dates contributed external predecessor bounds to a local schedule proposal.</summary>
public sealed class ProjectExternalScheduleSource {
    internal readonly ProjectDocument Document;
    internal ProjectExternalScheduleSource(string reference, ProjectDocument document, long revision) {
        Reference = reference; Document = document; ModelRevision = revision;
    }
    /// <summary>Caller-resolved reference recorded by the referring project.</summary>
    public string Reference { get; }
    /// <summary>Revision used to calculate the external dates. Applying the local proposal also checks this revision.</summary>
    public long ModelRevision { get; }
    internal void ValidateCurrent() {
        Document.EnsureNotDisposed();
        if (Document.Revision != ModelRevision || Document.HasActiveUpdate)
            throw new InvalidOperationException("An external project changed after this schedule was calculated.");
    }
}

internal sealed class ProjectExternalProjectContext {
    private readonly ProjectScheduleOptions _options;
    private readonly CancellationToken _token;
    private readonly HashSet<ProjectDocument> _active = new();
    private readonly Dictionary<(ProjectDocument Origin, string Reference), ProjectDocument> _resolved = new();
    private readonly Dictionary<ProjectDocument, ProjectScheduleResult> _calculated = new();
    private readonly Dictionary<ProjectDocument, long> _intervalCounts = new();
    private long _tasks;
    private long _intervals;
    internal ProjectExternalProjectContext(ProjectScheduleOptions options, CancellationToken token) { _options = options; _token = token; }
    internal void Enter(ProjectDocument document) {
        if (_active.Contains(document)) throw new InvalidDataException("External projects contain a dependency cycle.");
        if (_active.Count > _options.MaxExternalDepth) throw new InvalidDataException("External project dependencies exceed MaxExternalDepth.");
        _active.Add(document);
    }
    internal void Leave(ProjectDocument document) { _active.Remove(document); }
    internal bool WithinIntervalLimit(ProjectDocument document, long intervals) {
        _intervalCounts.TryGetValue(document, out long previous);
        _intervals = checked(_intervals - previous + intervals);
        _intervalCounts[document] = intervals;
        return _intervals <= _options.MaxIntervals;
    }
    internal (ProjectScheduleResult Schedule, ProjectTaskSchedule Task, string Reference) Resolve(ProjectDocument origin, ProjectDependency dependency) {
        _token.ThrowIfCancellationRequested();
        var resolver = _options.ExternalProjectResolver ?? throw new InvalidDataException("External dependencies require an explicit caller-controlled project resolver.");
        string combined = dependency.CrossProjectName ?? throw new InvalidDataException("The external dependency has no project reference.");
        int separator = Math.Max(combined.LastIndexOf('\\'), combined.LastIndexOf('/'));
        if (separator <= 0 || !int.TryParse(combined.Substring(separator + 1), NumberStyles.None, CultureInfo.InvariantCulture, out int displayId) || displayId < 1)
            throw new NotSupportedException("External predecessor references require the Project XML project-path/task-display-ID form.");
        string reference = combined.Substring(0, separator); var key = (origin, reference);
        if (!_resolved.TryGetValue(key, out var document)) {
            if (_resolved.Count >= _options.MaxExternalProjects) throw new InvalidDataException("External project resolution exceeds MaxExternalProjects.");
            document = resolver(origin, reference, _token) ?? throw new InvalidDataException("The caller did not resolve the external project reference.");
            _token.ThrowIfCancellationRequested(); document.EnsureNotDisposed();
            if (document.HasActiveUpdate) throw new InvalidOperationException("An external project has an unfinished update scope.");
            _resolved.Add(key, document);
        }
        if (_active.Contains(document)) throw new InvalidDataException("External projects contain a dependency cycle.");
        if (!_calculated.TryGetValue(document, out var schedule)) {
            // Reserve the source tasks before scheduling, including before resolving its children.
            // This bounds the aggregate graph rather than checking only after allocating schedules.
            foreach (var sourceTask in document.AllTasks) {
                _token.ThrowIfCancellationRequested();
                if (++_tasks > _options.MaxExternalTasks)
                    throw new InvalidDataException("The resolved external projects exceed MaxExternalTasks.");
            }
            schedule = new ProjectScheduler(document, _options, _token, externalContext: this).Calculate();
            schedule.Report.ThrowIfErrors();
            _calculated.Add(document, schedule);
        }
        if (schedule.ModelRevision != document.Revision || document.HasActiveUpdate) throw new InvalidOperationException("An external project changed during resolution.");
        var matches = document.AllTasks.Where(t => t.DisplayId == displayId).ToArray();
        if (matches.Length != 1) throw new InvalidDataException("The external task display ID does not identify exactly one task.");
        var task = schedule.Tasks.SingleOrDefault(t => t.TaskUid == matches[0].Uid);
        if (task == null || task.IsSummary) throw new NotSupportedException("External predecessors must resolve to calculated, active, non-summary tasks.");
        return (schedule, task, reference);
    }
}
