using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>A typed Microsoft Project document. XML editing preserves source data and never implicitly schedules tasks.</summary>
public sealed partial class ProjectDocument : IDisposable {
    private bool _disposed;
    internal bool Loading;
    private long _savedRevision;
    private int _batchDepth;
    private bool _batchChanged;
    internal bool HasPendingBatchChanges => _batchChanged;
    private DocumentAccessMode _accessMode = DocumentAccessMode.ReadWrite;
    private DocumentPersistenceMode _persistenceMode;
    private string? _path;
    private Stream? _associatedStream;
    internal ProjectXmlSource? Source;
    internal ProjectNativeSource? NativeSource;
    internal ProjectMpxSource? MpxSource;
    internal readonly Guid NativeIdentity = System.Guid.NewGuid();
    /// <summary>Native source generation and inert container inventory; null for XML/new documents.</summary>
    public ProjectNativeInfo? NativeInfo => NativeSource?.Info;
    internal byte[]? LastSavedBytes;
    internal readonly Dictionary<int, ProjectTask> TaskIndex = new Dictionary<int, ProjectTask>();
    internal readonly Dictionary<int, ProjectResource> ResourceIndex = new Dictionary<int, ProjectResource>();
    internal readonly Dictionary<int, ProjectCalendar> CalendarIndex = new Dictionary<int, ProjectCalendar>();
    internal readonly Dictionary<int, ProjectAssignment> AssignmentIndex = new Dictionary<int, ProjectAssignment>();
    // A packed Int64 hashes by XORing its halves; adjacent UIDs then collide heavily in dense schedules.
    internal readonly HashSet<(int First, int Second)> AssignmentPairs = new HashSet<(int, int)>();
    internal readonly HashSet<(int First, int Second)> DependencyPairs = new HashSet<(int, int)>();
    internal static (int First, int Second) PairKey(int first, int second) => (first, second);
    private int _taskUid = 1, _resourceUid = 1, _calendarUid = 1, _assignmentUid = 1;
    private readonly List<ProjectDiagnostic> _readDiagnostics = new List<ProjectDiagnostic>();
    internal int ReadDiagnosticLimit = int.MaxValue;

    private ProjectDocument() {
        Settings = new ProjectSettings(this);
        Tasks = new ProjectTaskCollection(this, null);
        Resources = new ProjectResourceCollection(this);
        Calendars = new ProjectCalendarCollection(this);
        Assignments = new ProjectAssignmentCollection(this);
        Dependencies = new ProjectDependencyCollection(this);
        CustomFields = new ProjectCollection<ProjectCustomFieldDefinition>(this, () => new ProjectCustomFieldDefinition(this));
    }

    private string? _name, _title, _author, _subject, _company, _manager;
    private Guid? _guid;
    private void SetMetadata<T>(ref T field, T value) { EnsureMutable(); if (EqualityComparer<T>.Default.Equals(field, value)) return; field = value; Touch(false); }
    /// <summary>Project name, separate from the associated file path.</summary>
    public string? Name { get => _name; set => SetMetadata(ref _name, value); }
    /// <summary>Document title.</summary>
    public string? Title { get => _title; set => SetMetadata(ref _title, value); }
    /// <summary>Document author. No machine identity is supplied automatically.</summary>
    public string? Author { get => _author; set => SetMetadata(ref _author, value); }
    /// <summary>Document subject.</summary>
    public string? Subject { get => _subject; set => SetMetadata(ref _subject, value); }
    /// <summary>Organization metadata.</summary>
    public string? Company { get => _company; set => SetMetadata(ref _company, value); }
    /// <summary>Project manager metadata.</summary>
    public string? Manager { get => _manager; set => SetMetadata(ref _manager, value); }
    /// <summary>Optional project GUID.</summary>
    public Guid? Guid { get => _guid; set => SetMetadata(ref _guid, value); }
    /// <summary>Document settings and project calendar reference.</summary>
    public ProjectSettings Settings { get; }
    /// <summary>Convenience access to the project's explicit calendar.</summary>
    public ProjectCalendar? Calendar { get => Settings.Calendar; set => Settings.Calendar = value; }
    /// <summary>Top-level tasks in outline order. Use AllTasks for a flat traversal.</summary>
    public ProjectTaskCollection Tasks { get; }
    /// <summary>All tasks in outline order, including project summaries and placeholders from source.</summary>
    public IEnumerable<ProjectTask> AllTasks => Traverse(Tasks);
    /// <summary>Resources.</summary>
    public ProjectResourceCollection Resources { get; }
    /// <summary>Base and derived calendars.</summary>
    public ProjectCalendarCollection Calendars { get; }
    /// <summary>Task-resource assignments.</summary>
    public ProjectAssignmentCollection Assignments { get; }
    /// <summary>Predecessor relationships.</summary>
    public ProjectDependencyCollection Dependencies { get; }
    /// <summary>Project-wide custom-field definitions.</summary>
    public ProjectCollection<ProjectCustomFieldDefinition> CustomFields { get; }
    /// <summary>Current mutation revision; no calculation is implied.</summary>
    public long Revision { get; private set; }
    /// <summary>True when edits have not been saved to an associated destination.</summary>
    public bool IsModified => Revision != _savedRevision || _batchChanged;
    /// <summary>True when schedule-affecting changes have been made without an explicit calculation.</summary>
    public bool IsScheduleStale { get; private set; }
    /// <summary>True after edits that can affect stored work/cost totals. Applying task dates does not recalculate those totals.</summary>
    public bool AreWorkCostTotalsStale { get; private set; }
    /// <summary>The input's SaveVersion, if declared. This is not an MPP write capability.</summary>
    public int? SourceSaveVersion => Source?.SaveVersion;
    /// <summary>Source XML namespace, or the native Project XML namespace for a new document.</summary>
    public string XmlNamespace => Source?.NamespaceName ?? ProjectXmlCodec.NamespaceName;
    /// <summary>Findings collected during input parsing, without exposing mutable XML.</summary>
    public IReadOnlyList<ProjectDiagnostic> ReadDiagnostics => _readDiagnostics.AsReadOnly();
    internal bool StructureChanged { get; private set; }

    internal void EnsureNotDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(ProjectDocument)); }
    internal void EnsureMutable() {
        EnsureNotDisposed();
        if (_accessMode == DocumentAccessMode.ReadOnly && !Loading) throw new InvalidOperationException("The document is read-only.");
    }
    internal void CheckMember(ProjectObject? value) {
        EnsureNotDisposed();
        if (value != null && (value.Document != this || !value.Attached)) throw new ArgumentException("The object must belong to this project and still be attached.");
    }
    internal void Touch(bool schedule, bool structure = false) {
        if (Loading) return;
        EnsureMutable();
        if (_batchDepth > 0) _batchChanged = true; else Revision++;
        IsScheduleStale |= schedule; AreWorkCostTotalsStale |= schedule; StructureChanged |= structure;
    }
    internal void AddReadDiagnostic(ProjectDiagnostic diagnostic) {
        if (_readDiagnostics.Count < ReadDiagnosticLimit) _readDiagnostics.Add(diagnostic);
        else if (_readDiagnostics.Count == ReadDiagnosticLimit) _readDiagnostics.Add(new ProjectDiagnostic("PROJECT_DIAGNOSTICS_TRUNCATED",
            ProjectDiagnosticSeverity.Warning, "Additional source diagnostics exceeded MaxDiagnostics.", "/Project"));
    }
    internal static IEnumerable<ProjectTask> Traverse(IEnumerable<ProjectTask> roots) {
        var stack = new Stack<IEnumerator<ProjectTask>>();
        stack.Push(roots.GetEnumerator());
        try {
            while (stack.Count != 0) {
                var current = stack.Peek();
                if (!current.MoveNext()) { current.Dispose(); stack.Pop(); continue; }
                var task = current.Current; yield return task;
                if (task.Children.Count != 0) stack.Push(task.Children.GetEnumerator());
            }
        } finally { while (stack.Count != 0) stack.Pop().Dispose(); }
    }

    /// <summary>Coalesces mutation revisions. This is not a rollback transaction and never calculates a schedule.</summary>
    public IDisposable BeginUpdate() { EnsureMutable(); _batchDepth++; return new UpdateScope(this); }
    private sealed class UpdateScope : IDisposable {
        private ProjectDocument? _document;
        internal UpdateScope(ProjectDocument document) { _document = document; }
        public void Dispose() {
            if (_document == null) return;
            var document = _document; _document = null;
            if (--document._batchDepth == 0 && document._batchChanged) { document.Revision++; document._batchChanged = false; }
        }
    }

    internal int NextTaskUid() => NextUid(TaskIndex, ref _taskUid);
    internal int NextResourceUid() => NextUid(ResourceIndex, ref _resourceUid);
    internal int NextCalendarUid() => NextUid(CalendarIndex, ref _calendarUid);
    internal int NextAssignmentUid() => NextUid(AssignmentIndex, ref _assignmentUid);
    private static int NextUid<T>(Dictionary<int, T> existing, ref int cursor) {
        while (existing.ContainsKey(cursor)) cursor = checked(cursor + 1);
        int result = cursor; cursor = checked(cursor + 1); return result;
    }
}
