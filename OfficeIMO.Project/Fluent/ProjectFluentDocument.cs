namespace OfficeIMO.Project.Fluent;

/// <summary>Fluent access to the same project model used by normal APIs.</summary>
public static class ProjectFluentExtensions {
    /// <summary>Begins a fluent configuration session with local, explicit task/resource aliases.</summary>
    public static ProjectFluentDocument AsFluent(this ProjectDocument document) => new ProjectFluentDocument(document);
}

/// <summary>Thin fluent configuration; End resolves forward aliases before creating relationships.</summary>
public sealed class ProjectFluentDocument {
    internal ProjectDocument Document { get; }
    private readonly Dictionary<string, ProjectTask> _tasks = new Dictionary<string, ProjectTask>(StringComparer.Ordinal);
    private readonly Dictionary<string, ProjectResource> _resources = new Dictionary<string, ProjectResource>(StringComparer.Ordinal);
    private readonly List<PendingLink> _links = new List<PendingLink>();
    private readonly List<PendingAssignment> _assignments = new List<PendingAssignment>();
    /// <summary>Wraps an existing document; no second model or scheduler is created.</summary>
    public ProjectFluentDocument(ProjectDocument document) {
        Document = document ?? throw new ArgumentNullException(nameof(document)); Document.EnsureMutable();
    }
    /// <summary>Configures metadata on the underlying document.</summary>
    public ProjectFluentDocument Info(Action<ProjectInfoBuilder> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        configure(new ProjectInfoBuilder(Document)); return this;
    }
    /// <summary>Sets a local project start and forward-scheduling input.</summary>
    public ProjectFluentDocument StartsOn(DateTime date) { Document.Settings.StartDate = date; Document.Settings.ScheduleFromStart = true; return this; }
    /// <summary>Creates and assigns an explicit standard working-week calendar.</summary>
    public ProjectFluentDocument StandardWorkingWeek(string name = "Standard") { Document.Calendar = Document.Calendars.AddStandardWorkingWeek(name); return this; }
    /// <summary>Configures an explicit calendar through the same typed model.</summary>
    public ProjectFluentDocument Calendar(string name, Action<ProjectCalendar> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        var calendar = Document.Calendars.Add(name); configure(calendar); Document.Calendar = calendar; return this;
    }
    /// <summary>Creates a resource and registers its session-local alias.</summary>
    public ProjectFluentDocument Resource(string alias, Action<ProjectResourceBuilder> configure) {
        CheckAlias(alias, _resources);
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        var resource = Document.Resources.AddWork(alias); _resources.Add(alias, resource);
        configure(new ProjectResourceBuilder(resource)); return this;
    }
    /// <summary>Registers an existing resource for later fluent assignments.</summary>
    public ProjectFluentDocument Resource(string alias, ProjectResource resource) {
        CheckAlias(alias, _resources); Document.CheckMember(resource);
        if (resource == null) throw new ArgumentNullException(nameof(resource));
        _resources.Add(alias, resource); return this;
    }
    /// <summary>Adds a top-level task.</summary>
    public ProjectFluentDocument Task(string alias, string name, Action<ProjectTaskBuilder>? configure = null) {
        AddTask(Document.Tasks, alias, name, false, configure); return this;
    }
    /// <summary>Adds and configures a top-level summary.</summary>
    public ProjectFluentDocument Summary(string alias, string name, Action<ProjectFluentTasks> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        var task = AddTask(Document.Tasks, alias, name, true, null);
        configure(new ProjectFluentTasks(this, task.Children)); return this;
    }
    /// <summary>Registers an existing task by stable UID and optionally edits it.</summary>
    public ProjectFluentDocument EditTask(string alias, int uid, Action<ProjectTaskBuilder>? configure = null) {
        CheckAlias(alias, _tasks);
        var task = Document.Tasks.GetByUid(uid); _tasks.Add(alias, task); configure?.Invoke(new ProjectTaskBuilder(this, task)); return this;
    }
    internal ProjectTask AddTask(ProjectTaskCollection tasks, string alias, string name, bool summary, Action<ProjectTaskBuilder>? configure) {
        CheckAlias(alias, _tasks);
        var task = summary ? tasks.AddSummary(name) : tasks.Add(name);
        _tasks.Add(alias, task); configure?.Invoke(new ProjectTaskBuilder(this, task)); return task;
    }
    internal void Link(ProjectTask successor, string predecessor, ProjectDependencyType type, ProjectDuration? lag) {
        if (string.IsNullOrWhiteSpace(predecessor)) throw new ArgumentException("A predecessor alias is required.", nameof(predecessor));
        _links.Add(new PendingLink(successor, predecessor, type, lag));
    }
    internal void Assign(ProjectTask task, string resource, ProjectUnits? units) {
        if (string.IsNullOrWhiteSpace(resource)) throw new ArgumentException("A resource alias is required.", nameof(resource));
        _assignments.Add(new PendingAssignment(task, resource, units));
    }

    /// <summary>Resolves pending aliases and returns the underlying document. No save or calculation occurs.</summary>
    /// <remarks>Scalar edits are immediate. Unresolved aliases create no pending relationships and can be corrected before retrying End.</remarks>
    public ProjectDocument End() {
        Document.EnsureMutable();
        var edges = new HashSet<string>(Document.Dependencies.Where(d => d.Predecessor != null).Select(d => d.Predecessor!.Uid + ":" + d.Successor.Uid), StringComparer.Ordinal);
        var assigned = new HashSet<string>(Document.Assignments.Where(a => a.Task != null && a.Resource != null).Select(a => a.Task!.Uid + ":" + a.Resource!.Uid), StringComparer.Ordinal);
        foreach (var item in _links) {
            if (!_tasks.TryGetValue(item.Alias, out var predecessor)) throw new InvalidOperationException("Unresolved task alias: " + item.Alias);
            Document.CheckMember(predecessor); Document.CheckMember(item.Successor);
            if (predecessor == item.Successor || !Enum.IsDefined(typeof(ProjectDependencyType), item.Type) || !edges.Add(predecessor.Uid + ":" + item.Successor.Uid))
                throw new InvalidOperationException("Invalid or duplicate fluent dependency: " + item.Alias);
        }
        foreach (var item in _assignments) {
            if (!_resources.TryGetValue(item.Alias, out var resource)) throw new InvalidOperationException("Unresolved resource alias: " + item.Alias);
            Document.CheckMember(resource); Document.CheckMember(item.Task);
            if (!assigned.Add(item.Task.Uid + ":" + resource.Uid)) throw new InvalidOperationException("Duplicate fluent assignment: " + item.Alias);
        }
        foreach (var item in _links) Document.Dependencies.Add(_tasks[item.Alias], item.Successor, item.Type).Lag = item.Lag;
        foreach (var item in _assignments) Document.Assignments.Add(item.Task, _resources[item.Alias], item.Units);
        _links.Clear(); _assignments.Clear(); return Document;
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
