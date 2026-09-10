namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    internal ProjectTask AddTask(ProjectTaskCollection collection, string name, bool summary) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        EnsureMutable(); CheckMember(collection.Parent);
        if (collection.Parent?.Uid == 0) throw new InvalidOperationException("The reserved project summary cannot own outline tasks. Add tasks to the document root collection.");
        var task = new ProjectTask(this, NextTaskUid()) { Name = name, Parent = collection.Parent, SourceSummary = summary };
        collection.Items.Add(task); TaskIndex.Add(task.Uid, task); Touch(true, true); return task;
    }

    internal void MoveTask(ProjectTask task, ProjectTask? parent, int? index) {
        EnsureMutable(); CheckMember(task); CheckMember(parent);
        if (task.Uid == 0 || parent?.Uid == 0) throw new InvalidOperationException("The reserved project summary cannot be moved or used as an outline parent.");
        for (var ancestor = parent; ancestor != null; ancestor = ancestor.Parent)
            if (ancestor == task) throw new ArgumentException("A task cannot be moved into its own subtree.");
        var source = task.Parent?.Children ?? Tasks;
        var destination = parent?.Children ?? Tasks;
        int countAfterRemoval = destination.Count - (destination == source ? 1 : 0);
        int position = index ?? countAfterRemoval;
        if (position < 0 || position > countAfterRemoval) throw new ArgumentOutOfRangeException(nameof(index));
        source.Items.Remove(task); destination.Items.Insert(position, task); task.Parent = parent;
        Touch(true, true);
    }

    internal void RemoveTask(ProjectTask task, ProjectRemovalMode mode) {
        EnsureMutable(); CheckMember(task); CheckRemovalMode(mode);
        var subtree = new HashSet<ProjectTask>(Traverse(new[] { task }));
        var links = Dependencies.Where(d => (d.Predecessor != null && subtree.Contains(d.Predecessor)) || subtree.Contains(d.Successor)).ToArray();
        var assignments = Assignments.Where(a => a.Task != null && subtree.Contains(a.Task)).ToArray();
        if (mode == ProjectRemovalMode.RejectIfReferenced && (task.Children.Count != 0 || links.Length != 0 || assignments.Length != 0))
            throw new InvalidOperationException("The task has children, dependencies, or assignments. Use explicit cascade removal.");
        foreach (var link in links) Dependencies.Remove(link);
        foreach (var assignment in assignments) Assignments.Remove(assignment);
        (task.Parent?.Children ?? Tasks).Items.Remove(task);
        foreach (var item in subtree) { item.Attached = false; TaskIndex.Remove(item.Uid); }
        Touch(true, true);
    }

    internal bool RemoveResource(ProjectResource resource, ProjectRemovalMode mode) {
        EnsureMutable(); CheckRemovalMode(mode);
        if (!Resources.Items.Contains(resource)) return false;
        var assignments = Assignments.Where(a => a.Resource == resource).ToArray();
        if (mode == ProjectRemovalMode.RejectIfReferenced && assignments.Length != 0)
            throw new InvalidOperationException("The resource has assignments. Use explicit cascade removal.");
        foreach (var assignment in assignments) Assignments.Remove(assignment);
        resource.Attached = false; Resources.Items.Remove(resource); ResourceIndex.Remove(resource.Uid);
        Touch(true, true); return true;
    }

    internal bool RemoveCalendar(ProjectCalendar calendar) {
        EnsureMutable();
        if (!Calendars.Items.Contains(calendar)) return false;
        if (Calendar == calendar || Calendars.Any(c => c.BaseCalendar == calendar) || AllTasks.Any(t => t.Calendar == calendar) || Resources.Any(r => r.Calendar == calendar))
            throw new InvalidOperationException("The calendar is still referenced.");
        calendar.Attached = false; Calendars.Items.Remove(calendar); CalendarIndex.Remove(calendar.Uid);
        Touch(true, true); return true;
    }

    private static void CheckRemovalMode(ProjectRemovalMode mode) {
        if (!Enum.IsDefined(typeof(ProjectRemovalMode), mode)) throw new ArgumentOutOfRangeException(nameof(mode));
    }

    /// <summary>Clones the current model through the selected format's validation and preservation contract, without an associated destination.</summary>
    public ProjectDocument Clone(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default, ProjectLoadOptions? loadOptions = null) {
        EnsureNotDisposed(); options ??= new ProjectSaveOptions(); options.Validate();
        loadOptions ??= new ProjectLoadOptions(); loadOptions.ValidateLimits();
        if (loadOptions.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) throw new ArgumentException("A clone has no associated destination.", nameof(loadOptions));
        var prepared = Serialize(options, cancellationToken);
        if (prepared.Bytes.Length > loadOptions.MaxInputBytes) throw new InvalidDataException("Project clone exceeds its input budget.");
        return ReadBytes((byte[])prepared.Bytes.Clone(), loadOptions, cancellationToken);
    }
}
