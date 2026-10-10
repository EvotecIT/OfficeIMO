using System.Collections;

namespace OfficeIMO.Access {
    /// <summary>An object identity belongs to a document and survives update rollback for pre-existing objects.</summary>
    public abstract class AccessNamedObject {
        internal AccessNamedObject(AccessDocument document, string name, Guid? identity = null) { Document = document; Name = ValidateName(name); Id = identity ?? Guid.NewGuid(); }
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
        public Guid Id { get; }
        /// <summary>Containing document identity.</summary>
        public Guid DocumentId => Document.Id;
        /// <summary>Object name. Rename semantics are introduced with qualified reference updates.</summary>
        public string Name { get; }
        /// <summary>Diagnostics specific to this modeled object.</summary>
        public IReadOnlyList<AccessDiagnostic> Diagnostics { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessDiagnostic>());
    }

    /// <summary>A named, document-owned collection. NotDecoded catalogs must not be interpreted as empty databases.</summary>
    public class AccessObjectCollection<T> : IReadOnlyList<T> where T : AccessNamedObject {
        internal readonly List<T> Items = new List<T>();
        internal readonly AccessDocument Document;
        internal AccessObjectCollection(AccessDocument document) { Document = document; }
        /// <summary>Whether this collection's native definitions are available. Application payload collections are qualified separately from tables.</summary>
        public AccessCatalogStatus CatalogStatus { get; internal set; } = AccessCatalogStatus.Modeled;
        /// <summary>Number of modeled objects; check CatalogStatus for a native source.</summary>
        public int Count { get { Document.EnsureNotDisposed(); return Items.Count; } }
        /// <summary>Object in model order.</summary>
        public T this[int index] { get { Document.EnsureNotDisposed(); return Items[index]; } }
        /// <summary>Object by an unambiguous case-insensitive name.</summary>
        public T this[string name] {
            get {
                Document.EnsureNotDisposed();
                if (CatalogStatus == AccessCatalogStatus.NotDecoded) throw new NotSupportedException("Native Access catalog lookup is not qualified. Inspect CatalogStatus and operation capabilities.");
                T? match = null;
                foreach (T item in Items) {
                    if (!StringComparer.OrdinalIgnoreCase.Equals(item.Name, name)) continue;
                    if (match != null) throw new InvalidOperationException($"The name '{name}' identifies more than one native object. Select its catalog type explicitly.");
                    match = item;
                }
                return match ?? throw new KeyNotFoundException($"No modeled object named '{name}'.");
            }
        }
        internal void AddItem(T item) {
            Document.EnsureMutable();
            if (Items.Any(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, item.Name))) throw new ArgumentException("An object with this name already exists.");
            Items.Add(item);
            Document.Changed(() => { Items.Remove(item); item.Attached = false; }, item.Id);
        }
        internal void AddNativeItem(T item) {
            if (Items.Any(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, item.Name))) throw new InvalidDataException("The native Access catalog contains duplicate or ambiguous names.");
            Items.Add(item);
        }
        /// <summary>Enumerates a stable snapshot of the modeled object references.</summary>
        public IEnumerator<T> GetEnumerator() { Document.EnsureNotDisposed(); return Items.ToArray().AsEnumerable().GetEnumerator(); }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>Modeled database tables.</summary>
    public sealed class AccessTableCollection : AccessObjectCollection<AccessTable> {
        internal AccessTableCollection(AccessDocument document) : base(document) { }
        /// <summary>Adds an in-memory table without creating a native catalog record.</summary>
        public AccessTable Add(string name) { AccessTable table = new AccessTable(Document, name); AddItem(table); return table; }
    }

    /// <summary>Inert saved query definitions; no SQL is executed.</summary>
    public sealed class AccessQueryCollection : AccessObjectCollection<AccessQueryDefinition> {
        internal AccessQueryCollection(AccessDocument document) : base(document) { }
        /// <summary>Adds an inert query definition. This slice does not parse, validate or execute Access SQL.</summary>
        public AccessQueryDefinition Add(string name, string sql) { AccessQueryDefinition query = new AccessQueryDefinition(Document, name, sql); AddItem(query); return query; }
    }

    /// <summary>An inert query definition whose text is never executed by the document library.</summary>
    public sealed class AccessQueryDefinition : AccessNamedObject {
        private readonly string? _sql;
        internal AccessQueryDefinition(AccessDocument document, string name, string sql) : this(document, name, sql ?? throw new ArgumentNullException(nameof(sql)), Array.Empty<AccessQueryRecord>()) { }
        internal AccessQueryDefinition(AccessDocument document, string name, string? sql, AccessQueryRecord[] records) : base(document, name) { _sql = sql; NativeRecords = Array.AsReadOnly(records); }
        /// <summary>Authored SQL or qualified normalized native SQL. Unqualified reconstruction throws; exact records remain available.</summary>
        public string Sql => _sql ?? throw new NotSupportedException("Native Access SQL reconstruction is unqualified for this query. Inspect NativeRecords and Diagnostics.");
        /// <summary>Whether Sql is available; native formatting is normalized rather than reproduced byte-for-byte.</summary>
        public bool HasSql => _sql != null;
        /// <summary>Inert native query records, including unsupported attributes and exact bytes.</summary>
        public IReadOnlyList<AccessQueryRecord> NativeRecords { get; }
        /// <summary>Persisted query object flags.</summary>
        public int NativeFlags { get; internal set; }
        /// <summary>Declared native parameters, without execution or supplied values.</summary>
        public IReadOnlyList<AccessQueryParameter> Parameters { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessQueryParameter>());
    }
}
