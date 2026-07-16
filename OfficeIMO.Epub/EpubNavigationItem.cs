namespace OfficeIMO.Epub;

/// <summary>Represents one ordered item in an EPUB table of contents, page list, or landmarks list.</summary>
public sealed class EpubNavigationItem {
    /// <summary>Package structure that supplied this item.</summary>
    public EpubNavigationSource Source { get; internal set; }

    /// <summary>Human-readable navigation label.</summary>
    public string Label { get; internal set; } = string.Empty;

    /// <summary>Original href or NCX src value.</summary>
    public string? Href { get; internal set; }

    /// <summary>Normalized archive path or absolute remote URI targeted by this item.</summary>
    public string? Target { get; internal set; }

    /// <summary>Decoded fragment identifier without the leading hash.</summary>
    public string? Fragment { get; internal set; }

    /// <summary>EPUB semantic type, NCX page type, or EPUB 2 guide type.</summary>
    public string? SemanticType { get; internal set; }

    /// <summary>NCX playOrder value when declared.</summary>
    public int? PlayOrder { get; internal set; }

    /// <summary>Whether the target is an absolute remote URI.</summary>
    public bool IsRemote { get; internal set; }

    /// <summary>Ordered child navigation items.</summary>
    public IReadOnlyList<EpubNavigationItem> Children { get; internal set; } = Array.Empty<EpubNavigationItem>();
}
