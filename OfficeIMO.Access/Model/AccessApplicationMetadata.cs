namespace OfficeIMO.Access;

/// <summary>An inert table data-macro definition with native provenance and exact XML source.</summary>
public sealed class AccessDataMacroInfo {
    internal AccessDataMacroInfo(AccessCatalogEntry source, string eventName, string xml, string[] statements) {
        CatalogEntry = source; Event = eventName; Xml = xml; Statements = Array.AsReadOnly(statements);
    }
    /// <summary>Source table catalog entry; its NativePayloads retain the exact MR2 macro map.</summary>
    public AccessCatalogEntry CatalogEntry { get; }
    /// <summary>Stored event/property name. Named/event macros are never invoked.</summary>
    public string Event { get; }
    /// <summary>Exact decoded macro XML. External entities and DTDs are prohibited during inspection.</summary>
    public string Xml { get; }
    /// <summary>Command element names in stored order, without expression/action evaluation.</summary>
    public IReadOnlyList<string> Statements { get; }
}

/// <summary>Native application resource metadata, separate from evaluating a form or report.</summary>
public sealed class AccessResourceInfo {
    internal AccessResourceInfo(int id, string? name, string? type, string? extension, AccessComplexValue? data, byte[] nativeRecord) {
        NativeId = id; Name = name; Type = type; Extension = extension; Data = data;
        NativeRecord = new AccessOpaqueValue(0, nativeRecord, "Exact resource row; embedded content and paths remain inert.");
    }
    /// <summary>Persisted resource row identity.</summary>
    public int NativeId { get; }
    /// <summary>Stored resource name.</summary>
    public string? Name { get; }
    /// <summary>Stored resource kind, without inferring rendering support.</summary>
    public string? Type { get; }
    /// <summary>Stored resource extension; no file is created or resolved.</summary>
    public string? Extension { get; }
    /// <summary>Lazy embedded attachments for the qualified structured resource layout.</summary>
    public AccessComplexValue? Data { get; }
    /// <summary>Exact row including unknown fields and native payload references.</summary>
    public AccessOpaqueValue NativeRecord { get; }
}
