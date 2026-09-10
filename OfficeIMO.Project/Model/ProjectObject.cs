namespace OfficeIMO.Project;

/// <summary>A mutable part of one project; mutations obey the owning document's access and disposal policy.</summary>
public abstract class ProjectObject {
    internal ProjectObject(ProjectDocument document) { Document = document; }
    internal ProjectDocument Document { get; }
    private bool _attached = true;
    internal ProjectObject? Owner { get; set; }
    internal bool Attached { get => _attached && (Owner?.Attached ?? true); set => _attached = value; }

    internal void EnsureAttached() {
        Document.EnsureMutable();
        if (!Attached) throw new InvalidOperationException("This object has been removed from its project.");
    }

    internal void Set<T>(ref T field, T value, bool schedule = false) {
        EnsureAttached();
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        Document.Touch(schedule);
    }

    internal void CheckReference(ProjectObject? value) {
        EnsureAttached();
        if (value != null && (value.Document != Document || !value.Attached))
            throw new ArgumentException("The referenced object must belong to this project and still be attached.");
    }
}

/// <summary>An entity whose stable UID is separate from its display position.</summary>
public abstract class ProjectEntity : ProjectObject {
    internal ProjectEntity(ProjectDocument document, int uid) : base(document) { Uid = uid; }
    /// <summary>Stable identity used by relationships; not renumbered during save.</summary>
    public int Uid { get; }
    private Guid? _guid;
    /// <summary>Optional source GUID; absent GUIDs remain absent.</summary>
    public Guid? Guid { get => _guid; set => Set(ref _guid, value); }
}

/// <summary>A named task, resource, or calendar with a stable identity.</summary>
public abstract class ProjectNamedEntity : ProjectEntity {
    internal ProjectNamedEntity(ProjectDocument document, int uid) : base(document, uid) { }
    private string? _name;
    /// <summary>The display name; duplicate names do not imply shared identity.</summary>
    public string? Name { get => _name; set => Set(ref _name, value); }
}

/// <summary>A bounded-by-document collection that creates items owned by the same project.</summary>
public class ProjectCollection<T> : IReadOnlyList<T> where T : ProjectObject {
    internal readonly List<T> Items = new List<T>();
    internal readonly ProjectDocument Document;
    private readonly Func<T> _factory;
    private readonly bool _schedule;
    private readonly ProjectObject? _owner;
    internal ProjectCollection(ProjectDocument document, Func<T> factory, bool schedule = false, ProjectObject? owner = null) {
        Document = document; _factory = factory; _schedule = schedule; _owner = owner;
    }
    /// <summary>Number of items in document order.</summary>
    public int Count => Items.Count;
    /// <summary>Returns an item by zero-based position.</summary>
    public T this[int index] => Items[index];
    /// <summary>Creates and appends an item owned by this document.</summary>
    public T Add() {
        Document.EnsureMutable();
        _owner?.EnsureAttached();
        var item = _factory();
        item.Owner = _owner;
        Items.Add(item);
        Document.Touch(_schedule, true);
        return item;
    }
    /// <summary>Removes an item; references to the removed object can no longer mutate it.</summary>
    public bool Remove(T item) {
        Document.EnsureMutable();
        _owner?.EnsureAttached();
        if (item == null) throw new ArgumentNullException(nameof(item));
        if (!Items.Contains(item)) return false;
        Items.Remove(item);
        item.Attached = false;
        Document.Touch(_schedule, true);
        return true;
    }
    /// <summary>Enumerates items in document order.</summary>
    public IEnumerator<T> GetEnumerator() => Items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
