namespace OfficeIMO.Project;

/// <summary>Immediate tasks in an outline, with lookup by stable UID across the document.</summary>
public sealed class ProjectTaskCollection : IReadOnlyList<ProjectTask> {
    internal readonly List<ProjectTask> Items = new List<ProjectTask>();
    internal ProjectTaskCollection(ProjectDocument document, ProjectTask? parent) { Document = document; Parent = parent; }
    internal ProjectDocument Document { get; }
    internal ProjectTask? Parent { get; }
    /// <summary>Number of immediate tasks.</summary>
    public int Count => Items.Count;
    /// <summary>Immediate task by zero-based position.</summary>
    public ProjectTask this[int index] => Items[index];
    /// <summary>Appends a task with a new stable UID.</summary>
    public ProjectTask Add(string name) => Document.AddTask(this, name, false);
    /// <summary>Appends a summary, which can contain child tasks.</summary>
    public ProjectTask AddSummary(string name) => Document.AddTask(this, name, true);
    /// <summary>Finds a task by UID in the whole document, not by row number.</summary>
    public ProjectTask GetByUid(int uid) => Document.TaskIndex.TryGetValue(uid, out var task) ? task : throw new KeyNotFoundException("Task UID " + uid + " does not exist.");
    /// <summary>Removes an immediate task with an explicit relationship policy.</summary>
    public bool Remove(ProjectTask task, ProjectRemovalMode mode = ProjectRemovalMode.RejectIfReferenced) {
        if (!Items.Contains(task)) return false;
        Document.RemoveTask(task, mode); return true;
    }
    /// <inheritdoc />
    public IEnumerator<ProjectTask> GetEnumerator() => Items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>Project resources with typed creation helpers.</summary>
public sealed class ProjectResourceCollection : IReadOnlyList<ProjectResource> {
    internal readonly List<ProjectResource> Items = new List<ProjectResource>();
    private readonly ProjectDocument _document;
    internal ProjectResourceCollection(ProjectDocument document) { _document = document; }
    /// <summary>Number of resources.</summary>
    public int Count => Items.Count;
    /// <summary>Resource by position.</summary>
    public ProjectResource this[int index] => Items[index];
    /// <summary>Creates a work resource.</summary>
    public ProjectResource AddWork(string name) => Add(name, ProjectResourceType.Work);
    /// <summary>Creates a material resource.</summary>
    public ProjectResource AddMaterial(string name) => Add(name, ProjectResourceType.Material);
    /// <summary>Creates a cost resource.</summary>
    public ProjectResource AddCost(string name) => Add(name, ProjectResourceType.Cost);
    private ProjectResource Add(string name, ProjectResourceType type) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        _document.EnsureMutable();
        var result = new ProjectResource(_document, _document.NextResourceUid()) { Name = name, Type = type };
        Items.Add(result); _document.ResourceIndex.Add(result.Uid, result);
        _document.Touch(true, true); return result;
    }
    /// <summary>Finds a resource by stable UID.</summary>
    public ProjectResource GetByUid(int uid) => _document.ResourceIndex.TryGetValue(uid, out var item) ? item : throw new KeyNotFoundException("Resource UID " + uid + " does not exist.");
    /// <summary>Removes a resource, rejecting assignments unless cascade removal is explicit.</summary>
    public bool Remove(ProjectResource resource, ProjectRemovalMode mode = ProjectRemovalMode.RejectIfReferenced) => _document.RemoveResource(resource, mode);
    /// <inheritdoc />
    public IEnumerator<ProjectResource> GetEnumerator() => Items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>Calendars owned by the project.</summary>
public sealed class ProjectCalendarCollection : IReadOnlyList<ProjectCalendar> {
    internal readonly List<ProjectCalendar> Items = new List<ProjectCalendar>();
    private readonly ProjectDocument _document;
    internal ProjectCalendarCollection(ProjectDocument document) { _document = document; }
    /// <summary>Number of calendars.</summary>
    public int Count => Items.Count;
    /// <summary>Calendar by position.</summary>
    public ProjectCalendar this[int index] => Items[index];
    /// <summary>Creates a calendar with no explicit working days.</summary>
    public ProjectCalendar Add(string name, ProjectCalendar? baseCalendar = null) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        _document.CheckMember(baseCalendar);
        _document.EnsureMutable();
        var result = new ProjectCalendar(_document, _document.NextCalendarUid()) { Name = name, BaseCalendar = baseCalendar, IsBaseCalendar = baseCalendar == null };
        Items.Add(result); _document.CalendarIndex.Add(result.Uid, result);
        _document.Touch(true, true); return result;
    }
    /// <summary>Creates an explicit Monday-Friday 08:00-12:00/13:00-17:00 working week.</summary>
    public ProjectCalendar AddStandardWorkingWeek(string name = "Standard") {
        var calendar = Add(name);
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek)))
            calendar.SetWorkingDay(day, day == DayOfWeek.Saturday || day == DayOfWeek.Sunday
                ? Array.Empty<ProjectWorkingTime>() : new[] { ProjectWorkingTime.Hours(8, 12), ProjectWorkingTime.Hours(13, 17) });
        return calendar;
    }
    /// <summary>Finds a calendar by stable UID.</summary>
    public ProjectCalendar GetByUid(int uid) => _document.CalendarIndex.TryGetValue(uid, out var item) ? item : throw new KeyNotFoundException("Calendar UID " + uid + " does not exist.");
    /// <summary>Removes an unreferenced calendar; inherited and assigned calendars cannot be removed.</summary>
    public bool Remove(ProjectCalendar calendar) => _document.RemoveCalendar(calendar);
    /// <inheritdoc />
    public IEnumerator<ProjectCalendar> GetEnumerator() => Items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
