namespace OfficeIMO.Markdown;

/// <summary>
/// Controls whether the renderer emits HTML asset tags or only returns an asset manifest for host-side merging.
/// </summary>
public enum MarkdownAssetEmitMode {
    /// <summary>Emit tags normally (inline/link), plus include the manifest.</summary>
    Emit,
    /// <summary>Do not emit tags; only return the manifest via <see cref="MarkdownHtmlParts.Assets"/>.</summary>
    ManifestOnly
}
