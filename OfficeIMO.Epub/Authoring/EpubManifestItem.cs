namespace OfficeIMO.Epub;

/// <summary>A manifest declaration backed by the writable package XML; unknown attributes are retained.</summary>
public sealed class EpubManifestItem {
    private readonly XElement _element;
    private readonly string _opfPath;
    private readonly EpubPublication? _publication;
    internal EpubManifestItem(XElement element, string opfPath, EpubPublication? publication = null) { _element = element; _opfPath = opfPath; _publication = publication; }
    private void Set(XName name, string? value) {
        if (_publication != null) _publication.SetDeclarationAttribute(_element, name, value);
        else _element.SetAttributeValue(name, value);
    }
    /// <summary>Package-local resource identifier.</summary>
    public string Id => (string?)_element.Attribute("id") ?? string.Empty;
    /// <summary>Original OPF-relative resource URL.</summary>
    public string Href => (string?)_element.Attribute("href") ?? string.Empty;
    /// <summary>Resolved reference using the canonical EPUB URL rules.</summary>
    public EpubReference Reference => EpubReference.Resolve(_opfPath, Href);
    /// <summary>Declared MIME media type.</summary>
    public string MediaType { get => (string?)_element.Attribute("media-type") ?? string.Empty; set => Set("media-type", value); }
    /// <summary>Space-separated package properties.</summary>
    public string? Properties { get => (string?)_element.Attribute("properties"); set => Set("properties", value); }
    /// <summary>Fallback manifest item id, when present.</summary>
    public string? FallbackId { get => (string?)_element.Attribute("fallback"); set => Set("fallback", value); }
    internal string? FallbackStyleId => (string?)_element.Attribute("fallback-style");
    /// <summary>Associated media-overlay manifest item id.</summary>
    public string? MediaOverlayId { get => (string?)_element.Attribute("media-overlay"); set => Set("media-overlay", value); }
}

/// <summary>One reading position, independent of manifest resource identity.</summary>
public sealed class EpubSpineItem {
    private readonly XElement _element;
    private readonly EpubPublication? _publication;
    internal EpubSpineItem(XElement element, EpubPublication? publication = null) { _element = element; _publication = publication; }
    private void Set(XName name, string? value) {
        if (_publication != null) _publication.SetDeclarationAttribute(_element, name, value);
        else _element.SetAttributeValue(name, value);
    }
    /// <summary>Manifest item selected at this position.</summary>
    public string ManifestId => (string?)_element.Attribute("idref") ?? string.Empty;
    /// <summary>Whether the item participates in the primary reading order.</summary>
    public bool IsLinear { get => (string?)_element.Attribute("linear") != "no"; set => Set("linear", value ? null : "no"); }
    /// <summary>Space-separated itemref properties, including layout/page-side declarations.</summary>
    public string? Properties { get => (string?)_element.Attribute("properties"); set => Set("properties", value); }
}

/// <summary>One authored navigation node. Targets are container-relative paths with optional fragments.</summary>
public sealed class EpubNavigationEntry {
    /// <summary>Creates a labelled navigation node.</summary>
    public EpubNavigationEntry(string label, string target, IEnumerable<EpubNavigationEntry>? children = null, string? semanticType = null) {
        Label = label ?? throw new ArgumentNullException(nameof(label));
        Target = target ?? throw new ArgumentNullException(nameof(target));
        Children = Array.AsReadOnly((children ?? Array.Empty<EpubNavigationEntry>()).ToArray());
        SemanticType = semanticType;
    }
    /// <summary>Visible label.</summary>
    public string Label { get; }
    /// <summary>Target interpreted from the container root, not the navigation directory.</summary>
    public string Target { get; }
    /// <summary>Semantic type for a landmark or page-list entry.</summary>
    public string? SemanticType { get; }
    /// <summary>Ordered child entries.</summary>
    public IReadOnlyList<EpubNavigationEntry> Children { get; }
}
