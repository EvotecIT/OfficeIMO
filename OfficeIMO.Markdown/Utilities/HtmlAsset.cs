namespace OfficeIMO.Markdown;

/// <summary>Represents a CSS/JS asset with a stable id for deduplication.</summary>
public sealed class HtmlAsset {
    /// <summary>Stable identifier used for deduplication, e.g. "prism-core" or "prism-lang:csharp".</summary>
    public string Id { get; }
    /// <summary>Asset type (CSS or JS).</summary>
    public HtmlAssetKind Kind { get; }
    /// <summary>URL to the asset when linked; null when inlined.</summary>
    public string? Href { get; }
    /// <summary>Inline content when not linked; null when referenced via Href.</summary>
    public string? Inline { get; }
    /// <summary>Optional media attribute for CSS links (e.g., "(prefers-color-scheme: dark)").</summary>
    public string? Media { get; set; }
    /// <summary>Creates a new asset descriptor.</summary>
    public HtmlAsset(string id, HtmlAssetKind kind, string? href, string? inline) { Id = id; Kind = kind; Href = href; Inline = inline; }
}

