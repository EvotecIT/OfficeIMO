namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private readonly Dictionary<string, ProjectTask> _taskAliases = new Dictionary<string, ProjectTask>(StringComparer.Ordinal);
    private readonly Dictionary<string, ProjectResource> _resourceAliases = new Dictionary<string, ProjectResource>(StringComparer.Ordinal);
    private readonly List<PendingLink> _links = new List<PendingLink>();
    private readonly List<PendingAssignment> _assignments = new List<PendingAssignment>();

    /// <summary>Sets the local start date and forward-scheduling input without running the scheduler.</summary>
    public ProjectDocument StartsOn(DateTime date) { Settings.StartDate = date; Settings.ScheduleFromStart = true; return this; }

    /// <summary>Creates and assigns an explicit standard working-week calendar.</summary>
    public ProjectDocument StandardWorkingWeek(string name = "Standard") { Calendar = Calendars.AddStandardWorkingWeek(name); return this; }

    internal void RegisterTaskAlias(string alias, ProjectTask task) {
        EnsureMutable(); CheckAlias(alias, _taskAliases); CheckMember(task); _taskAliases.Add(alias, task);
    }
    internal void RegisterResourceAlias(string alias, ProjectResource resource) {
        EnsureMutable(); CheckAlias(alias, _resourceAliases); CheckMember(resource); _resourceAliases.Add(alias, resource);
    }
    internal void QueueDependency(ProjectTask successor, string predecessor, ProjectDependencyType type, ProjectDuration? lag) {
        EnsureMutable();
        if (string.IsNullOrWhiteSpace(predecessor)) throw new ArgumentException("A predecessor alias is required.", nameof(predecessor));
        _links.Add(new PendingLink(successor, predecessor, type, lag));
    }
    internal void QueueAssignment(ProjectTask task, string resource, ProjectUnits? units) {
        EnsureMutable();
        if (string.IsNullOrWhiteSpace(resource)) throw new ArgumentException("A resource alias is required.", nameof(resource));
        _assignments.Add(new PendingAssignment(task, resource, units));
    }

    /// <summary>Resolves pending aliases and returns the underlying document. No save or calculation occurs.</summary>
    /// <remarks>Scalar edits are immediate. Unresolved aliases create no pending relationships and can be corrected before retrying End.</remarks>
    public ProjectDocument End() {
        EnsureMutable();
        var edges = new HashSet<string>(Dependencies.Where(d => d.Predecessor != null).Select(d => d.Predecessor!.Uid + ":" + d.Successor.Uid), StringComparer.Ordinal);
        var assigned = new HashSet<string>(Assignments.Where(a => a.Task != null && a.Resource != null).Select(a => a.Task!.Uid + ":" + a.Resource!.Uid), StringComparer.Ordinal);
        foreach (var item in _links) {
            if (!_taskAliases.TryGetValue(item.Alias, out var predecessor)) throw new InvalidOperationException("Unresolved task alias: " + item.Alias);
            CheckMember(predecessor); CheckMember(item.Successor);
            if (predecessor == item.Successor || !Enum.IsDefined(typeof(ProjectDependencyType), item.Type) || !edges.Add(predecessor.Uid + ":" + item.Successor.Uid))
                throw new InvalidOperationException("Invalid or duplicate deferred dependency: " + item.Alias);
        }
        foreach (var item in _assignments) {
            if (!_resourceAliases.TryGetValue(item.Alias, out var resource)) throw new InvalidOperationException("Unresolved resource alias: " + item.Alias);
            CheckMember(resource); CheckMember(item.Task);
            if (!assigned.Add(item.Task.Uid + ":" + resource.Uid)) throw new InvalidOperationException("Duplicate deferred assignment: " + item.Alias);
        }
        foreach (var item in _links) Dependencies.Add(_taskAliases[item.Alias], item.Successor, item.Type).Lag = item.Lag;
        foreach (var item in _assignments) Assignments.Add(item.Task, _resourceAliases[item.Alias], item.Units);
        _links.Clear(); _assignments.Clear(); return this;
    }
    private static void CheckAlias<T>(string alias, Dictionary<string, T> aliases) {
        if (string.IsNullOrWhiteSpace(alias)) throw new ArgumentException("An explicit alias is required.", nameof(alias));
        if (aliases.ContainsKey(alias)) throw new ArgumentException("Duplicate alias: " + alias, nameof(alias));
    }
    private sealed class PendingLink {
        internal PendingLink(ProjectTask successor, string alias, ProjectDependencyType type, ProjectDuration? lag) { Successor = successor; Alias = alias; Type = type; Lag = lag; }
        internal ProjectTask Successor { get; }
        internal string Alias { get; }
        internal ProjectDependencyType Type { get; }
        internal ProjectDuration? Lag { get; }
    }
    private sealed class PendingAssignment {
        internal PendingAssignment(ProjectTask task, string alias, ProjectUnits? units) { Task = task; Alias = alias; Units = units; }
        internal ProjectTask Task { get; }
        internal string Alias { get; }
        internal ProjectUnits? Units { get; }
    }
}
