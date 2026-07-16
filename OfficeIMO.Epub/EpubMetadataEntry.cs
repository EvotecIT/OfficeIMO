namespace OfficeIMO.Epub;

/// <summary>Represents one ordered OPF metadata declaration.</summary>
public sealed class EpubMetadataEntry {
    /// <summary>Metadata entry kind.</summary>
    public EpubMetadataKind Kind { get; internal set; }

    /// <summary>Element local name, such as title, creator, meta, or link.</summary>
    public string Name { get; internal set; } = string.Empty;

    /// <summary>Element namespace URI.</summary>
    public string NamespaceUri { get; internal set; } = string.Empty;

    /// <summary>Normalized text, content, or href value.</summary>
    public string Value { get; internal set; } = string.Empty;

    /// <summary>Element id used by refinements or package identity.</summary>
    public string? Id { get; internal set; }

    /// <summary>EPUB 3 property name.</summary>
    public string? Property { get; internal set; }

    /// <summary>EPUB 3 refinement target.</summary>
    public string? Refines { get; internal set; }

    /// <summary>Declared value scheme.</summary>
    public string? Scheme { get; internal set; }

    /// <summary>xml:lang value.</summary>
    public string? Language { get; internal set; }

    /// <summary>EPUB 2 meta name.</summary>
    public string? LegacyName { get; internal set; }

    /// <summary>Creator/contributor role.</summary>
    public string? Role { get; internal set; }

    /// <summary>File-as sorting value.</summary>
    public string? FileAs { get; internal set; }

    /// <summary>Date event classification.</summary>
    public string? Event { get; internal set; }

    /// <summary>Linked metadata href.</summary>
    public string? Href { get; internal set; }

    /// <summary>Linked metadata relation tokens.</summary>
    public string? Rel { get; internal set; }

    /// <summary>Linked metadata media type.</summary>
    public string? MediaType { get; internal set; }
}
