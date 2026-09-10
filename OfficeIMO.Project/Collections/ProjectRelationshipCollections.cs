namespace OfficeIMO.Project;

/// <summary>Task links owned by a project.</summary>
public sealed class ProjectDependencyCollection : IReadOnlyList<ProjectDependency> {
    internal readonly List<ProjectDependency> Items = new List<ProjectDependency>();
    private readonly ProjectDocument _document;
    internal ProjectDependencyCollection(ProjectDocument document) { _document = document; }
    /// <summary>Number of relationships.</summary>
    public int Count => Items.Count;
    /// <summary>Relationship by position.</summary>
    public ProjectDependency this[int index] => Items[index];
    /// <summary>Adds a local task dependency after checking identities and duplicate links.</summary>
    public ProjectDependency Add(ProjectTask predecessor, ProjectTask successor, ProjectDependencyType type = ProjectDependencyType.FinishToStart) {
        if (predecessor == null) throw new ArgumentNullException(nameof(predecessor));
        if (successor == null) throw new ArgumentNullException(nameof(successor));
        _document.EnsureMutable(); _document.CheckMember(predecessor); _document.CheckMember(successor);
        if (predecessor == successor) throw new ArgumentException("A task cannot depend on itself.");
        if (!Enum.IsDefined(typeof(ProjectDependencyType), type)) throw new ArgumentOutOfRangeException(nameof(type));
        long key = ProjectDocument.PairKey(predecessor.Uid, successor.Uid);
        if (_document.DependencyPairs.Contains(key)) throw new ArgumentException("The dependency already exists.");
        var result = new ProjectDependency(_document) { Predecessor = predecessor, Successor = successor, SourcePredecessorUid = predecessor.Uid, Type = type };
        Items.Add(result); _document.DependencyPairs.Add(key); _document.Touch(true, true); return result;
    }
    /// <summary>Removes a relationship without removing either task.</summary>
    public bool Remove(ProjectDependency dependency) {
        _document.EnsureMutable();
        if (!Items.Contains(dependency)) return false;
        Items.Remove(dependency); dependency.Attached = false;
        if (dependency.Predecessor != null) _document.DependencyPairs.Remove(ProjectDocument.PairKey(dependency.Predecessor.Uid, dependency.Successor.Uid));
        _document.Touch(true, true); return true;
    }
    /// <inheritdoc />
    public IEnumerator<ProjectDependency> GetEnumerator() => Items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>Resource assignments owned by the project.</summary>
public sealed class ProjectAssignmentCollection : IReadOnlyList<ProjectAssignment> {
    internal readonly List<ProjectAssignment> Items = new List<ProjectAssignment>();
    private readonly ProjectDocument _document;
    internal ProjectAssignmentCollection(ProjectDocument document) { _document = document; }
    /// <summary>Number of assignments.</summary>
    public int Count => Items.Count;
    /// <summary>Assignment by position.</summary>
    public ProjectAssignment this[int index] => Items[index];
    /// <summary>Adds an assignment between entities from this document.</summary>
    public ProjectAssignment Add(ProjectTask task, ProjectResource resource, ProjectUnits? units = null) {
        if (task == null) throw new ArgumentNullException(nameof(task));
        if (resource == null) throw new ArgumentNullException(nameof(resource));
        _document.EnsureMutable(); _document.CheckMember(task); _document.CheckMember(resource);
        long key = ProjectDocument.PairKey(task.Uid, resource.Uid);
        if (_document.AssignmentPairs.Contains(key)) throw new ArgumentException("This resource is already assigned to the task.");
        var result = new ProjectAssignment(_document, _document.NextAssignmentUid()) {
            Task = task, Resource = resource, SourceTaskUid = task.Uid, SourceResourceUid = resource.Uid,
            Units = units ?? ProjectUnits.Percent(100)
        };
        Items.Add(result); _document.AssignmentIndex.Add(result.Uid, result); _document.AssignmentPairs.Add(key); _document.Touch(true, true); return result;
    }
    /// <summary>Finds an assignment by stable UID.</summary>
    public ProjectAssignment GetByUid(int uid) => _document.AssignmentIndex.TryGetValue(uid, out var assignment) ? assignment : throw new KeyNotFoundException("Assignment UID " + uid + " does not exist.");
    /// <summary>Removes an assignment while retaining its task and resource.</summary>
    public bool Remove(ProjectAssignment assignment) {
        _document.EnsureMutable();
        if (!Items.Contains(assignment)) return false;
        Items.Remove(assignment); assignment.Attached = false;
        _document.AssignmentIndex.Remove(assignment.Uid);
        if (assignment.Task != null && assignment.Resource != null) _document.AssignmentPairs.Remove(ProjectDocument.PairKey(assignment.Task.Uid, assignment.Resource.Uid));
        _document.Touch(true, true); return true;
    }
    /// <inheritdoc />
    public IEnumerator<ProjectAssignment> GetEnumerator() => Items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
