namespace OfficeIMO.Access;

/// <summary>A typed Access model with explicit operation qualification. Instances are not thread-safe.</summary>
public sealed partial class AccessDocument : IDisposable {
    private bool _disposed;
    private int _readers;
    private AccessUpdateScope? _update;
    private readonly List<Action> _undo = new List<Action>();
    private readonly List<AccessChange> _changes = new List<AccessChange>();
    private string? _path;
    private Stream? _destination;
    private long _inputLimit;
    private int _pageLimit;
    internal AccessNativeDatabase? NativeDatabase;
    private AccessNativeWriter? _creationPlan;
    private long _creationPlanLimit;
    internal DateTime CreatedAt { get; } = DateTime.Now;

    private AccessDocument(AccessFileFormat format, DocumentAccessMode accessMode, AccessInspection? inspection) {
        Format = format; AccessMode = accessMode; Inspection = inspection;
        Profile = inspection?.Profile ?? (format == AccessFileFormat.Mdb ? AccessFormatProfile.Jet4 : AccessFormatProfile.Ace12);
        CatalogStatus = inspection == null ? AccessCatalogStatus.Modeled : AccessCatalogStatus.NotDecoded;
        Tables = new AccessTableCollection(this); Queries = new AccessQueryCollection(this);
        Relationships = new AccessRelationshipCollection(this);
        Forms = new AccessObjectCollection<AccessApplicationObject>(this);
        Reports = new AccessObjectCollection<AccessApplicationObject>(this);
        Macros = new AccessObjectCollection<AccessApplicationObject>(this);
        SystemTables = new AccessObjectCollection<AccessTable>(this);
        Catalog = new AccessObjectCollection<AccessCatalogEntry>(this);
        VbaProject = new AccessVbaProjectInfo(CatalogStatus);
        Tables.CatalogStatus = Queries.CatalogStatus = Relationships.CatalogStatus = CatalogStatus;
        SystemTables.CatalogStatus = Catalog.CatalogStatus = CatalogStatus;
        Forms.CatalogStatus = Reports.CatalogStatus = Macros.CatalogStatus = CatalogStatus;
    }
    /// <summary>Unique model identity.</summary>
    public Guid Id { get; } = Guid.NewGuid();
    /// <summary>Mutation revision. Rollback restores its starting revision.</summary>
    public long Revision { get; private set; }
    /// <summary>Whether the model has edits. No native persistence is implied.</summary>
    public bool IsModified => Revision != 0;
    /// <summary>Whether an edit scope must finish before reads or assessment.</summary>
    public bool HasActiveUpdate => _update != null;
    /// <summary>Whether the loaded catalog is decoded or the objects are a new model.</summary>
    public AccessCatalogStatus CatalogStatus { get; internal set; }
    /// <summary>Selected or detected file family.</summary>
    public AccessFileFormat Format { get; }
    /// <summary>Selected or detected physical generation.</summary>
    public AccessFormatProfile Profile { get; }
    /// <summary>Document mutation policy.</summary>
    public DocumentAccessMode AccessMode { get; }
    /// <summary>Native creation and unchanged snapshot preservation require an explicit Save.</summary>
    public DocumentPersistenceMode PersistenceMode => DocumentPersistenceMode.Explicit;
    /// <summary>Bounded native header evidence; null for a newly created model.</summary>
    public AccessInspection? Inspection { get; }
    /// <summary>Modeled tables; native catalog decoding is a separate capability.</summary>
    public AccessTableCollection Tables { get; }
    /// <summary>System and hidden table definitions, separate from ordinary user tables. Opening them never executes database logic.</summary>
    public AccessObjectCollection<AccessTable> SystemTables { get; }
    /// <summary>Inert native catalog entries, including system and unknown objects. Catalog inventory does not decode application payloads.</summary>
    public AccessObjectCollection<AccessCatalogEntry> Catalog { get; }
    /// <summary>Document-level decoding limits and unsupported-content diagnostics.</summary>
    public IReadOnlyList<AccessDiagnostic> Diagnostics { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessDiagnostic>());
    /// <summary>Persisted database properties, including the AccessVersion format marker. Expressions remain inert.</summary>
    public IReadOnlyDictionary<string, object?> Properties { get; internal set; } = new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?>());
    /// <summary>Exact database property map, including unqualified chunks.</summary>
    public AccessOpaqueValue? NativeProperties { get; internal set; }
    /// <summary>Stored legacy database code page. Jet4/ACE text values use Unicode; this marker does not authorize lossy decoding.</summary>
    public int? CodePage { get; internal set; }
    /// <summary>Stored collation identifier. Reading metadata does not implement engine sorting or query evaluation.</summary>
    public int? SortOrder { get; internal set; }
    /// <summary>Inert modeled query definitions.</summary>
    public AccessQueryCollection Queries { get; }
    /// <summary>Typed modeled relationships.</summary>
    public AccessRelationshipCollection Relationships { get; }
    /// <summary>Form metadata boundary. Check CatalogStatus before interpreting absence.</summary>
    public AccessObjectCollection<AccessApplicationObject> Forms { get; }
    /// <summary>Report metadata boundary. Check CatalogStatus before interpreting absence.</summary>
    public AccessObjectCollection<AccessApplicationObject> Reports { get; }
    /// <summary>Action macro metadata boundary, separate from VBA.</summary>
    public AccessObjectCollection<AccessApplicationObject> Macros { get; }
    /// <summary>Inert VBA inventory availability.</summary>
    public AccessVbaProjectInfo VbaProject { get; internal set; }
    /// <summary>Inert application storage streams with native paths, including unknown and compiled content.</summary>
    public IReadOnlyList<AccessStorageStream> ApplicationStreams { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessStorageStream>());
    /// <summary>Inert observed query, designer and relationship references. This is a partial dependency map, never an expression engine.</summary>
    public IReadOnlyList<AccessDependency> Dependencies { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessDependency>());
    /// <summary>Qualified native table data-macro XML definitions. Actions and expressions are retained without execution.</summary>
    public IReadOnlyList<AccessDataMacroInfo> DataMacros { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessDataMacroInfo>());
    /// <summary>Native resource metadata; embedded payloads use the existing lazy attachment reader.</summary>
    public IReadOnlyList<AccessResourceInfo> Resources { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessResourceInfo>());
    /// <summary>Committed or pending model mutations; rollback removes mutations from its scope.</summary>
    public IReadOnlyList<AccessChange> ChangeJournal { get { EnsureNotDisposed(); return _changes.AsReadOnly(); } }
    /// <summary>Operation boundaries for this document state, including unsupported native operations.</summary>
    public IReadOnlyList<AccessOperationCapability> Capabilities {
        get {
            EnsureNotDisposed();
            return Array.AsReadOnly(AccessCapabilities.Operations.Select(operation => {
                bool enabled = operation.IsSupported;
                if (operation.Operation == "model.edit") enabled &= CatalogStatus == AccessCatalogStatus.Modeled && AccessMode == DocumentAccessMode.ReadWrite;
                if (operation.Operation == "model.rows.read") enabled &= CatalogStatus != AccessCatalogStatus.NotDecoded;
                if (operation.Operation == "native.catalog.read" || operation.Operation == "native.rows.read" || operation.Operation == "native.query.records.read") enabled &= CatalogStatus == AccessCatalogStatus.Decoded;
                if (operation.Operation == "native.create") enabled &= CatalogStatus == AccessCatalogStatus.Modeled;
                if (operation.Operation == "native.preserve") enabled &= Inspection != null;
                if (operation.Operation == "application.objects.read") enabled &= Forms.CatalogStatus == AccessCatalogStatus.Decoded;
                if (operation.Operation == "vba.inspect") enabled &= VbaProject.CatalogStatus == AccessCatalogStatus.Decoded;
                return new AccessOperationCapability(operation.Operation, enabled, operation.Boundary);
            }).ToArray());
        }
    }

    internal void EnsureNotDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(AccessDocument)); }
    internal void EnsureMutable() {
        EnsureNotDisposed();
        if (AccessMode == DocumentAccessMode.ReadOnly) throw new InvalidOperationException("This Access document is read-only.");
        if (CatalogStatus != AccessCatalogStatus.Modeled) throw new NotSupportedException("Native Access editing is not qualified. Loaded native documents are immutable.");
        if (_readers != 0) throw new InvalidOperationException("Close Access data readers before modifying the document.");
        if (Revision == long.MaxValue) throw new InvalidOperationException("The model revision limit was reached.");
    }
    internal void Changed(Action undo, Guid? objectId = null, string operation = "object.add") {
        _creationPlan = null;
        Revision++;
        _changes.Add(new AccessChange(Revision, objectId ?? Id, operation));
        if (_update != null) _undo.Add(undo);
    }
    internal void AcquireReader() {
        EnsureNotDisposed();
        if (HasActiveUpdate) throw new InvalidOperationException("Commit or roll back the Access update before opening a reader.");
        _readers = checked(_readers + 1);
    }
    internal void ReleaseReader() { _readers--; }
    /// <summary>Begins a model edit scope. Disposal rolls back unless Commit succeeds; database transactions are separate.</summary>
    public AccessUpdateScope BeginUpdate() {
        EnsureMutable();
        if (_update != null) throw new InvalidOperationException("Nested Access update scopes are not supported.");
        return _update = new AccessUpdateScope(this, Revision);
    }
    internal void FinishUpdate(AccessUpdateScope scope, long revision, bool commit) {
        if (_update != scope) throw new InvalidOperationException("This Access update scope is no longer active.");
        if (!commit) { for (int i = _undo.Count - 1; i >= 0; i--) _undo[i](); _changes.RemoveAll(x => x.Revision > revision); Revision = revision; }
        _undo.Clear(); _update = null;
    }
    /// <summary>Discards pending model edits and releases the document. Caller-owned streams remain open.</summary>
    public void Dispose() {
        if (_disposed) return;
        _update?.Dispose(); _disposed = true; _destination = null; _creationPlan = null; NativeDatabase?.Dispose(); NativeDatabase = null;
    }
}

/// <summary>An explicit model edit scope with rollback on disposal.</summary>
public sealed class AccessUpdateScope : IDisposable {
    private readonly AccessDocument _document;
    private readonly long _revision;
    private bool _finished;
    internal AccessUpdateScope(AccessDocument document, long revision) { _document = document; _revision = revision; }
    /// <summary>Accepts model edits. It does not write a native database or execute a database transaction.</summary>
    public void Commit() {
        _document.EnsureNotDisposed();
        if (_finished) throw new InvalidOperationException("This Access update scope is complete.");
        _document.FinishUpdate(this, _revision, true); _finished = true;
    }
    /// <summary>Rolls back uncommitted edits while retaining identities of pre-existing objects.</summary>
    public void Dispose() {
        if (_finished) return;
        _document.FinishUpdate(this, _revision, false); _finished = true;
    }
}
