namespace OfficeIMO.Markdown;

internal static class UrlOriginPolicy {
    internal static bool IsAllowedHttpLink(HtmlOptions? o, string? url)
        => IsAllowedHttpUrl(o, url, forImages: false);

    internal static bool IsAllowedHttpImage(HtmlOptions? o, string? url) {
        if (o == null) return true;
        var u = (url ?? string.Empty).Trim();
        if (u.Length == 0) return true;

        if (o.BlockExternalHttpImages) {
            // Treat "external" relative to BaseUri when available; otherwise, any absolute HTTP(S) image is external.
            // Relative images (including "/path") are not blocked by this option.
            if (TryGetAbsoluteHttpUri(u, o.BaseUri, out var abs) && abs != null && IsHttpScheme(abs.Scheme)) {
                var baseUri = o.BaseUri;
                if (baseUri != null && baseUri.IsAbsoluteUri && IsHttpScheme(baseUri.Scheme)) {
                    if (!IsSameOrigin(baseUri, abs)) return false;
                } else {
                    return false;
                }
            }
        }

        return IsAllowedHttpUrl(o, u, forImages: true);
    }

    private static bool IsAllowedHttpUrl(HtmlOptions? o, string? url, bool forImages) {
        if (o == null) return true;
        var u = (url ?? string.Empty).Trim();
        if (u.Length == 0) return true;
        if (u.StartsWith("#", StringComparison.Ordinal)) return true; // fragment-only

        // Host allowlist (absolute HTTP(S) only).
        var allowHosts = forImages ? o.AllowedHttpImageHosts : o.AllowedHttpLinkHosts;
        if (allowHosts != null && allowHosts.Count > 0) {
            if (TryGetAbsoluteHttpUri(u, o.BaseUri, out var absForHost) && absForHost != null && IsHttpScheme(absForHost.Scheme)) {
                if (!HostAllowList.IsAllowed(absForHost.Host, allowHosts)) return false;
            }
        }

        bool restrict = forImages ? o.RestrictHttpImagesToBaseOrigin : o.RestrictHttpLinksToBaseOrigin;
        if (!restrict) return true;

        var baseUri = o.BaseUri;
        if (baseUri == null || !baseUri.IsAbsoluteUri) return true;
        if (!IsHttpScheme(baseUri.Scheme)) return true; // don't attempt "origin" semantics for non-http(s) bases

        // Relative URLs are considered within base origin.
        if (!TryGetAbsoluteHttpUri(u, baseUri, out var abs) || abs == null) return true;

        if (!IsHttpScheme(abs.Scheme)) return true; // mailto, etc.

        return IsSameOrigin(baseUri, abs);
    }

    private static bool TryGetAbsoluteHttpUri(string u, Uri? baseUri, out Uri? abs) {
        abs = null;
        if (u == null) return false;
        if (u.Trim().Length == 0) return false;

        // Protocol-relative URLs. Assume base scheme when known; fall back to https.
        if (u.StartsWith("//", StringComparison.Ordinal)) {
            var scheme = (baseUri != null && baseUri.IsAbsoluteUri) ? baseUri.Scheme : "https";
            return Uri.TryCreate(scheme + ":" + u, UriKind.Absolute, out abs) && abs != null;
        }

        return Uri.TryCreate(u, UriKind.Absolute, out abs) && abs != null;
    }

    private static bool IsAbsoluteExternalHttp(string u)
        => u.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
           || u.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
           || u.StartsWith("//", StringComparison.OrdinalIgnoreCase);

    private static bool IsHttpScheme(string? scheme)
        => "http".Equals(scheme, StringComparison.OrdinalIgnoreCase)
           || "https".Equals(scheme, StringComparison.OrdinalIgnoreCase);

    private static bool IsSameOrigin(Uri a, Uri b) {
        if (!string.Equals(a.Scheme, b.Scheme, StringComparison.OrdinalIgnoreCase)) return false;
        if (!string.Equals(a.Host, b.Host, StringComparison.OrdinalIgnoreCase)) return false;
        return GetEffectivePort(a) == GetEffectivePort(b);
    }

    private static int GetEffectivePort(Uri u) {
        if (!u.IsDefaultPort) return u.Port;
        if ("http".Equals(u.Scheme, StringComparison.OrdinalIgnoreCase)) return 80;
        if ("https".Equals(u.Scheme, StringComparison.OrdinalIgnoreCase)) return 443;
        return u.Port;
    }
}
