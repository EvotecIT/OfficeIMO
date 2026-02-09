namespace OfficeIMO.Markdown;

internal static class UrlOriginPolicy {
    internal static bool IsAllowedHttpLink(HtmlOptions? o, string? url)
        => IsAllowedHttpUrlByBaseOrigin(o, url, forImages: false);

    internal static bool IsAllowedHttpImage(HtmlOptions? o, string? url) {
        if (o == null) return true;
        var u = (url ?? string.Empty).Trim();
        if (u.Length == 0) return true;

        if (o.BlockExternalHttpImages) {
            if (IsAbsoluteExternalHttp(u)) return false;
        }

        return IsAllowedHttpUrlByBaseOrigin(o, u, forImages: true);
    }

    private static bool IsAllowedHttpUrlByBaseOrigin(HtmlOptions? o, string? url, bool forImages) {
        if (o == null) return true;
        var u = (url ?? string.Empty).Trim();
        if (u.Length == 0) return true;
        if (u.StartsWith("#", StringComparison.Ordinal)) return true; // fragment-only

        bool restrict = forImages ? o.RestrictHttpImagesToBaseOrigin : o.RestrictHttpLinksToBaseOrigin;
        if (!restrict) return true;

        var baseUri = o.BaseUri;
        if (baseUri == null || !baseUri.IsAbsoluteUri) return true;
        if (!IsHttpScheme(baseUri.Scheme)) return true; // don't attempt "origin" semantics for non-http(s) bases

        // Relative URLs are considered within base origin.
        if (!Uri.TryCreate(u, UriKind.Absolute, out var abs)) {
            // Protocol-relative URLs ("//host/path") are not absolute URIs per Uri.TryCreate; treat as disallowed when restricting.
            if (u.StartsWith("//", StringComparison.Ordinal)) return false;
            return true;
        }

        if (!IsHttpScheme(abs.Scheme)) return true; // mailto, etc.

        return IsSameOrigin(baseUri, abs);
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

