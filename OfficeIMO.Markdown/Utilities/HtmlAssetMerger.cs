namespace OfficeIMO.Markdown;

/// <summary>
/// Host-facing helper to merge multiple HtmlAsset manifests into deduplicated CSS/JS tags and inline blobs.
/// </summary>
public static class HtmlAssetMerger {
    /// <summary>
    /// Deduplicate by Id and build strings for head link/script tags and inline CSS/JS payloads.
    /// </summary>
    /// <param name="manifests">Asset manifests to merge in order.</param>
    /// <param name="options">Optional HTML rendering policy used for asset attribute encoding.</param>
    public static (string headLinks, string inlineCss, string inlineJs) Build(IEnumerable<IEnumerable<HtmlAsset>> manifests, HtmlOptions? options = null) {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var head = new StringBuilder();
        var css = new StringBuilder();
        var js = new StringBuilder();
        foreach (var manifest in manifests) {
            foreach (var a in manifest) {
                if (!seen.Add(a.Id)) continue;
                if (a.Kind == HtmlAssetKind.Css) {
                    if (!string.IsNullOrEmpty(a.Href)) {
                        var media = string.IsNullOrEmpty(a.Media) ? string.Empty : $" media=\"{HtmlTextEncoder.Encode(a.Media, options)}\"";
                        head.Append($"<link rel=\"stylesheet\" data-asset-id=\"{HtmlTextEncoder.Encode(a.Id, options)}\" href=\"{HtmlAttributeUrlEncoder.Encode(a.Href, options)}\"{media}>\n");
                    } else if (!string.IsNullOrEmpty(a.Inline)) {
                        if (css.Length > 0) css.Append('\n');
                        css.Append(a.Inline);
                    }
                } else {
                    if (!string.IsNullOrEmpty(a.Href)) {
                        head.Append($"<script data-asset-id=\"{HtmlTextEncoder.Encode(a.Id, options)}\" src=\"{HtmlAttributeUrlEncoder.Encode(a.Href, options)}\"></script>\n");
                    } else if (!string.IsNullOrEmpty(a.Inline)) {
                        if (js.Length > 0) js.Append('\n');
                        js.Append(a.Inline);
                    }
                }
            }
        }
        return (head.ToString(), css.ToString(), js.ToString());
    }
}

