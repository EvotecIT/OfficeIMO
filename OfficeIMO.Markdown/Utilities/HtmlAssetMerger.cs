using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown.Utilities;

/// <summary>
/// Host-facing helper to merge multiple HtmlAsset manifests into deduplicated CSS/JS tags and inline blobs.
/// </summary>
public static class HtmlAssetMerger {
    /// <summary>
    /// Deduplicate by Id and build strings for head link/script tags and inline CSS/JS payloads.
    /// </summary>
    public static (string headLinks, string inlineCss, string inlineJs) Build(IEnumerable<IEnumerable<HtmlAsset>> manifests) {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var head = new StringBuilder();
        var css = new StringBuilder();
        var js = new StringBuilder();
        foreach (var manifest in manifests) {
            foreach (var a in manifest) {
                if (!seen.Add(a.Id)) continue;
                if (a.Kind == HtmlAssetKind.Css) {
                    if (!string.IsNullOrEmpty(a.Href)) {
                        var media = string.IsNullOrEmpty(a.Media) ? string.Empty : $" media=\"{System.Net.WebUtility.HtmlEncode(a.Media)}\"";
                        head.Append($"<link rel=\"stylesheet\" data-asset-id=\"{System.Net.WebUtility.HtmlEncode(a.Id)}\" href=\"{System.Net.WebUtility.HtmlEncode(a.Href)}\"{media}>\n");
                    } else if (!string.IsNullOrEmpty(a.Inline)) {
                        if (css.Length > 0) css.Append('\n');
                        css.Append(a.Inline);
                    }
                } else {
                    if (!string.IsNullOrEmpty(a.Href)) {
                        head.Append($"<script data-asset-id=\"{System.Net.WebUtility.HtmlEncode(a.Id)}\" src=\"{System.Net.WebUtility.HtmlEncode(a.Href)}\"></script>\n");
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

