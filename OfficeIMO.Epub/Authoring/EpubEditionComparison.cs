namespace OfficeIMO.Epub;

/// <summary>Observable changes between publication editions.</summary>
[Flags]
public enum EpubEditionChangeKind {
    /// <summary>No difference.</summary>
    None = 0,
    /// <summary>A manifest resource or retained entry was added.</summary>
    Added = 1,
    /// <summary>A manifest resource or retained entry was removed.</summary>
    Removed = 2,
    /// <summary>A resource with the same manifest ID changed its resolved location.</summary>
    Location = 4,
    /// <summary>Manifest attributes or child declarations changed.</summary>
    Declaration = 8,
    /// <summary>XML element, attribute or node content changed.</summary>
    Xml = 16,
    /// <summary>XHTML body text or other XML root text changed, without rendering or whitespace normalization.</summary>
    Text = 32,
    /// <summary>Bytes changed while parsed XML stayed equal.</summary>
    Serialization = 64,
    /// <summary>Non-XML or encrypted payload bytes changed; their semantics were not interpreted.</summary>
    Binary = 128
}

/// <summary>A changed resource, matched by manifest ID, or an unmanifested entry matched by path.</summary>
public sealed class EpubEditionResourceChange {
    internal EpubEditionResourceChange(string? id, string? before, string? after, EpubEditionChangeKind kind) {
        ManifestId = id; PreviousPath = before; CurrentPath = after; Kind = kind;
    }
    /// <summary>Manifest identity, or null for an unmanifested retained entry.</summary>
    public string? ManifestId { get; }
    /// <summary>Original resolved URL or retained entry path.</summary>
    public string? PreviousPath { get; }
    /// <summary>Revised resolved URL or retained entry path.</summary>
    public string? CurrentPath { get; }
    /// <summary>Combined observed differences.</summary>
    public EpubEditionChangeKind Kind { get; }
}

/// <summary>Read-only edition differences. No content is modified or rendered.</summary>
public sealed class EpubEditionComparison {
    internal EpubEditionComparison(bool metadata, bool spine, bool package, IEnumerable<EpubEditionResourceChange> resources, IEnumerable<EpubEditionTextChange> textChanges) {
        MetadataChanged = metadata; ReadingOrderChanged = spine; PackageStructureChanged = package;
        Resources = Array.AsReadOnly(resources.ToArray());
        TextChanges = Array.AsReadOnly(textChanges.ToArray());
    }
    /// <summary>Package metadata XML changed, including refinements, excluding the unrefined dcterms:modified write timestamp.</summary>
    public bool MetadataChanged { get; }
    /// <summary>Spine order, repeated positions, attributes or reading declarations changed.</summary>
    public bool ReadingOrderChanged { get; }
    /// <summary>Package path, root attributes or non-metadata/non-manifest/non-spine content changed.</summary>
    public bool PackageStructureChanged { get; }
    /// <summary>Changed resources and unmanifested entries in deterministic identity order.</summary>
    public IReadOnlyList<EpubEditionResourceChange> Resources { get; }
    /// <summary>Changed XHTML text blocks and whole-body text, with bounded excerpts.</summary>
    public IReadOnlyList<EpubEditionTextChange> TextChanges { get; }
    /// <summary>Whether any compared package, resource or retained-entry difference was found.</summary>
    public bool HasChanges => MetadataChanged || ReadingOrderChanged || PackageStructureChanged || Resources.Count != 0;
}
