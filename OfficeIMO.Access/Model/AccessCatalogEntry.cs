namespace OfficeIMO.Access;

/// <summary>Inert native catalog inventory. An entry does not certify typed application-object decoding or preservation.</summary>
public sealed class AccessCatalogEntry : AccessNamedObject {
    internal AccessCatalogEntry(AccessDocument document, string name, int nativeType, int flags) : base(document, name) { NativeType = nativeType; Flags = flags; }
    /// <summary>Persisted catalog type, including unknown types.</summary>
    public int NativeType { get; }
    /// <summary>Persisted catalog flags. These are inspection evidence, not mutable physical identifiers.</summary>
    public int Flags { get; }
    /// <summary>Whether the entry is marked hidden or system.</summary>
    public bool IsSystem => (Flags & unchecked((int)0x80000002)) != 0;
    /// <summary>Exact catalog record, retained for explicit inspection rather than proof of native save preservation.</summary>
    public AccessOpaqueValue? NativeRecord { get; internal set; }
    /// <summary>Persisted owner/security identifier bytes, inspected without authentication or a workgroup lookup.</summary>
    public AccessOpaqueValue? Owner { get; internal set; }
}

/// <summary>Linked-table metadata with credentials redacted. This object never resolves files, networks or providers.</summary>
public sealed class AccessLinkedTableInfo {
    internal AccessLinkedTableInfo(string? source, string? foreignTable, string? connection) { Source = source; ForeignTableName = foreignTable; Connection = connection; }
    /// <summary>Stored source name or path, inspected without opening it.</summary>
    public string? Source { get; }
    /// <summary>Stored external table name.</summary>
    public string? ForeignTableName { get; }
    /// <summary>Stored connection metadata with credential values redacted.</summary>
    public string? Connection { get; }
}
