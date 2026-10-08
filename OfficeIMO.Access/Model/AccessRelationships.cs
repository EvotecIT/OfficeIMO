namespace OfficeIMO.Access;

/// <summary>Typed table index definitions. This slice does not enforce database constraints or allocate index pages.</summary>
public sealed class AccessIndexCollection : AccessObjectCollection<AccessIndex> {
    private readonly AccessTable _table;
    internal AccessIndexCollection(AccessTable table) : base(table.Document) { _table = table; }
    /// <summary>Adds a primary-key definition over named fields.</summary>
    public AccessIndex AddPrimaryKey(string name, params string[] columns) {
        _table.EnsureAttached(); Document.EnsureMutable();
        if (Items.Any(i => i.IsPrimaryKey)) throw new InvalidOperationException("A table already has a primary key definition.");
        if (columns == null || columns.Length == 0 || columns.Length > 10) throw new ArgumentException("An index requires 1–10 fields.", nameof(columns));
        AccessColumn[] fields = columns.Select(c => _table.Columns[c]).ToArray();
        if (fields.Distinct().Count() != fields.Length) throw new ArgumentException("An index cannot repeat a field.", nameof(columns));
        var index = new AccessIndex(_table, name, fields); AddItem(index); return index;
    }
}

/// <summary>Stable typed index definition, independent of physical index storage.</summary>
public sealed class AccessIndex : AccessNamedObject {
    internal AccessIndex(AccessTable table, string name, AccessColumn[] columns) : base(table.Document, name) { Table = table; Columns = Array.AsReadOnly(columns); }
    /// <summary>Owning table.</summary>
    public AccessTable Table { get; }
    /// <summary>Indexed fields in declared order.</summary>
    public IReadOnlyList<AccessColumn> Columns { get; }
    /// <summary>Whether this definition is a primary key.</summary>
    public bool IsPrimaryKey => true;
}

/// <summary>Modeled relationships with column identities rather than name-only references.</summary>
public sealed class AccessRelationshipCollection : AccessObjectCollection<AccessRelationship> {
    internal AccessRelationshipCollection(AccessDocument document) : base(document) { }
    /// <summary>Adds a single-field relationship. Runtime referential integrity is not executed by the model.</summary>
    public AccessRelationship Add(string name, AccessColumn parent, AccessColumn child) {
        Document.EnsureMutable();
        if (parent == null || child == null) throw new ArgumentNullException(parent == null ? nameof(parent) : nameof(child));
        parent.EnsureAttached(); child.EnsureAttached();
        if (parent.Document != Document || child.Document != Document) throw new ArgumentException("Relationship fields must belong to this document.");
        AccessDataType parentType = parent.DataType == AccessDataType.AutoNumber ? AccessDataType.Int32 : parent.DataType;
        AccessDataType childType = child.DataType == AccessDataType.AutoNumber ? AccessDataType.Int32 : child.DataType;
        if (parentType != childType) throw new ArgumentException("Relationship field types must agree.");
        var relationship = new AccessRelationship(Document, name, parent, child); AddItem(relationship); return relationship;
    }
}

/// <summary>An inert relationship definition.</summary>
public sealed class AccessRelationship : AccessNamedObject {
    internal AccessRelationship(AccessDocument document, string name, AccessColumn parent, AccessColumn child) : base(document, name) { Parent = parent; Child = child; }
    /// <summary>Referenced field.</summary>
    public AccessColumn Parent { get; }
    /// <summary>Referencing field.</summary>
    public AccessColumn Child { get; }
}
