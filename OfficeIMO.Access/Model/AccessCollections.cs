using System.Collections;

namespace OfficeIMO.Access;

/// <summary>An object identity belongs to a document and survives update rollback for pre-existing objects.</summary>
public abstract class AccessNamedObject {
    internal AccessNamedObject(AccessDocument document, string name) { Document = document; Name = ValidateName(name); }
    internal AccessDocument Document { get; }
    internal bool Attached = true;
    internal void EnsureAttached() {
        Document.EnsureNotDisposed();
        if (!Attached) throw new InvalidOperationException("This object was removed by update rollback.");
    }
    internal static string ValidateName(string name) {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 64 || name.IndexOf('\0') >= 0) throw new ArgumentException("Access object names must contain 1–64 non-null characters.", nameof(name));
        return name;
    }
    /// <summary>Stable model identity; it is separate from native page numbers and names.</summary>
    public Guid Id { get; } = Guid.NewGuid();
    /// <summary>Containing document identity.</summary>
    public Guid DocumentId => Document.Id;
    /// <summary>Object name. Rename semantics are introduced with qualified reference updates.</summary>
    public string Name { get; }
    /// <summary>Diagnostics specific to this modeled object.</summary>
    public IReadOnlyList<AccessDiagnostic> Diagnostics { get; } = Array.AsReadOnly(Array.Empty<AccessDiagnostic>());
}

/// <summary>A named, document-owned collection. NotDecoded catalogs must not be interpreted as empty databases.</summary>
public class AccessObjectCollection<T> : IReadOnlyList<T> where T : AccessNamedObject {
    internal readonly List<T> Items = new List<T>();
    internal readonly AccessDocument Document;
    internal AccessObjectCollection(AccessDocument document) { Document = document; }
    /// <summary>Number of modeled objects; check CatalogStatus for a native source.</summary>
    public int Count { get { Document.EnsureNotDisposed(); return Items.Count; } }
    /// <summary>Object in model order.</summary>
    public T this[int index] { get { Document.EnsureNotDisposed(); return Items[index]; } }
    /// <summary>Object by an unambiguous case-insensitive name.</summary>
    public T this[string name] { get { Document.EnsureNotDisposed(); if (Document.CatalogStatus == AccessCatalogStatus.NotDecoded) throw new NotSupportedException("Native Access catalog lookup is not qualified. Inspect CatalogStatus and operation capabilities."); return Items.FirstOrDefault(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, name)) ?? throw new KeyNotFoundException($"No modeled object named '{name}'."); } }
    internal void AddItem(T item) {
        Document.EnsureMutable();
        if (Items.Any(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, item.Name))) throw new ArgumentException("An object with this name already exists.");
        Items.Add(item);
        Document.Changed(() => { Items.Remove(item); item.Attached = false; });
    }
    /// <summary>Enumerates a stable snapshot of the modeled object references.</summary>
    public IEnumerator<T> GetEnumerator() { Document.EnsureNotDisposed(); return Items.ToArray().AsEnumerable().GetEnumerator(); }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>Modeled database tables.</summary>
public sealed class AccessTableCollection : AccessObjectCollection<AccessTable> {
    internal AccessTableCollection(AccessDocument document) : base(document) { }
    /// <summary>Adds an in-memory table without creating a native catalog record.</summary>
    public AccessTable Add(string name) { var table = new AccessTable(Document, name); AddItem(table); return table; }
}

/// <summary>Inert saved query definitions; no SQL is executed.</summary>
public sealed class AccessQueryCollection : AccessObjectCollection<AccessQueryDefinition> {
    internal AccessQueryCollection(AccessDocument document) : base(document) { }
    /// <summary>Adds an inert query definition. This slice does not parse, validate or execute Access SQL.</summary>
    public AccessQueryDefinition Add(string name, string sql) { var query = new AccessQueryDefinition(Document, name, sql); AddItem(query); return query; }
}

/// <summary>An inert query definition whose text is never executed by the document library.</summary>
public sealed class AccessQueryDefinition : AccessNamedObject {
    internal AccessQueryDefinition(AccessDocument document, string name, string sql) : base(document, name) { Sql = sql ?? throw new ArgumentNullException(nameof(sql)); }
    /// <summary>Access SQL definition. Catalog decoding and typed parameters require the query codec.</summary>
    public string Sql { get; }
}

/// <summary>Metadata boundary for an Access form, report or action macro.</summary>
public sealed class AccessApplicationObject : AccessNamedObject {
    internal AccessApplicationObject(AccessDocument document, string name) : base(document, name) { }
}

/// <summary>VBA inventory availability. No module or event procedure is executed.</summary>
public sealed class AccessVbaProjectInfo {
    internal AccessVbaProjectInfo(AccessCatalogStatus status) { CatalogStatus = status; }
    /// <summary>NotDecoded means modules and signatures have not been inspected, rather than being absent.</summary>
    public AccessCatalogStatus CatalogStatus { get; }
    /// <summary>Known module names; an empty list is conclusive only for a new modeled document.</summary>
    public IReadOnlyList<string> ModuleNames { get; } = Array.AsReadOnly(Array.Empty<string>());
}
