namespace OfficeIMO.Markdown;

internal static class UrlOriginPolicy {
    internal static bool IsAllowedHttpLink(HtmlOptions? o, string? url) {
        if (!IsSafeLinkUrl(url)) return false;
        return IsAllowedHttpUrl(o, url, forImages: false);
    }

    internal static bool IsAllowedHttpImage(HtmlOptions? o, string? url) {
        if (!IsSafeImageUrl(url)) return false;
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

    internal static string FilterAllowedImageSrcSet(HtmlOptions? options, string? srcSet) {
        if (string.IsNullOrWhiteSpace(srcSet)) return string.Empty;

        var allowed = new System.Collections.Generic.List<string>();
        foreach (SrcSetCandidate candidate in SrcSetParser.Parse(srcSet)) {
            if (!IsAllowedHttpImage(options, candidate.Url) || !IsValidSrcSetDescriptor(candidate.Descriptor)) {
                continue;
            }

            allowed.Add(candidate.Descriptor.Length == 0
                ? candidate.Url
                : candidate.Url + " " + candidate.Descriptor);
        }

        return string.Join(", ", allowed);
    }

    private static bool IsAllowedHttpUrl(HtmlOptions? o, string? url, bool forImages) {
        if (o == null) return true;
        var u = (url ?? string.Empty).Trim();
        if (u.Length == 0) return true;
        if (u.StartsWith("#", StringComparison.Ordinal)) return true; // fragment-only

        // Host allowlist (absolute HTTP(S) only).
        var allowHosts = forImages ? o.AllowedHttpImageHosts : o.AllowedHttpLinkHosts;
        bool restrict = forImages ? o.RestrictHttpImagesToBaseOrigin : o.RestrictHttpLinksToBaseOrigin;
        if ((restrict || allowHosts != null && allowHosts.Count > 0)
            && TryGetScheme(u, out string? explicitScheme)
            && (explicitScheme == "http" || explicitScheme == "https")
            && !IsWellFormedAbsoluteHttpUrl(u, explicitScheme)) {
            return false;
        }

        if (allowHosts != null && allowHosts.Count > 0) {
            if (TryGetAbsoluteHttpUri(u, o.BaseUri, out var absForHost) && absForHost != null && IsHttpScheme(absForHost.Scheme)) {
                if (!HostAllowList.IsAllowed(absForHost.Host, allowHosts)) return false;
            }
        }

        if (!restrict) return true;

        var baseUri = o.BaseUri;
        if (baseUri == null || !baseUri.IsAbsoluteUri) return true;
        if (!IsHttpScheme(baseUri.Scheme)) return true; // don't attempt "origin" semantics for non-http(s) bases

        // Relative URLs are considered within base origin.
        if (!TryGetAbsoluteHttpUri(u, baseUri, out var abs) || abs == null) return true;

        if (!IsHttpScheme(abs.Scheme)) return true; // mailto, etc.

        return IsSameOrigin(baseUri, abs);
    }

    private static bool IsSafeLinkUrl(string? url) {
        string value = (url ?? string.Empty).Trim();
        if (value.Length == 0 || value.StartsWith("#", StringComparison.Ordinal)) return true;
        if (!TryGetScheme(value, out string? scheme)) return true;

        return scheme == "http"
            || scheme == "https"
            || scheme == "mailto"
            || scheme == "tel";
    }

    private static bool IsSafeImageUrl(string? url) {
        string value = (url ?? string.Empty).Trim();
        if (value.Length == 0 || value.StartsWith("#", StringComparison.Ordinal)) return true;
        if (!TryGetScheme(value, out string? scheme)) return true;

        if (scheme == "http" || scheme == "https") {
            return IsWellFormedAbsoluteHttpUrl(value, scheme);
        }

        if (scheme == "cid") return true;
        return scheme == "data" && IsSafeRasterDataImage(value);
    }

    private static bool IsWellFormedAbsoluteHttpUrl(string value, string expectedScheme) {
        return Uri.TryCreate(value, UriKind.Absolute, out Uri? uri)
            && uri != null
            && string.Equals(uri.Scheme, expectedScheme, StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(uri.Host);
    }

    private static bool TryGetScheme(string value, out string? scheme) {
        scheme = null;
        int colon = value.IndexOf(':');
        if (colon <= 0) return false;

        int slash = value.IndexOfAny(new[] { '/', '?', '#' });
        if (slash >= 0 && slash < colon) return false;

        var builder = new System.Text.StringBuilder(colon);
        for (int i = 0; i < colon; i++) {
            char current = value[i];
            if (char.IsWhiteSpace(current) || char.IsControl(current)) continue;
            builder.Append(char.ToLowerInvariant(current));
        }

        if (builder.Length == 0 || !char.IsLetter(builder[0])) {
            scheme = string.Empty;
            return true;
        }

        for (int i = 1; i < builder.Length; i++) {
            char current = builder[i];
            if (!char.IsLetterOrDigit(current) && current != '+' && current != '-' && current != '.') {
                scheme = string.Empty;
                return true;
            }
        }

        scheme = builder.ToString();
        return true;
    }

    private static bool IsSafeRasterDataImage(string value) {
        int comma = value.IndexOf(',');
        if (comma <= 5) return false;

        string metadata = value.Substring(5, comma - 5).Trim();
        int separator = metadata.IndexOf(';');
        string mediaType = (separator >= 0 ? metadata.Substring(0, separator) : metadata).Trim();
        string parameters = separator >= 0 ? metadata.Substring(separator + 1) : string.Empty;
        bool base64 = parameters.Split(';').Any(parameter => string.Equals(parameter.Trim(), "base64", StringComparison.OrdinalIgnoreCase));
        if (!base64) return false;

        return mediaType.Equals("image/png", StringComparison.OrdinalIgnoreCase)
            || mediaType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase)
            || mediaType.Equals("image/gif", StringComparison.OrdinalIgnoreCase)
            || mediaType.Equals("image/webp", StringComparison.OrdinalIgnoreCase)
            || mediaType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase)
            || mediaType.Equals("image/tiff", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsValidSrcSetDescriptor(string descriptor) {
        if (string.IsNullOrEmpty(descriptor)) return true;
        if (descriptor.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) >= 0 || descriptor.Length < 2) return false;

        char suffix = descriptor[descriptor.Length - 1];
        string number = descriptor.Substring(0, descriptor.Length - 1);
        if (suffix == 'w') {
            return int.TryParse(number, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int width)
                && width > 0;
        }

        return suffix == 'x'
            && double.TryParse(number, System.Globalization.NumberStyles.AllowDecimalPoint, System.Globalization.CultureInfo.InvariantCulture, out double density)
            && density > 0
            && !double.IsInfinity(density)
            && !double.IsNaN(density);
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
