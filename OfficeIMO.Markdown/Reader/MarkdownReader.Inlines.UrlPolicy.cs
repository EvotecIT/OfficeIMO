namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static string? ResolveImageUrl(string url, MarkdownLoadOptions? options) =>
        options?.ImageSourceResolver is { } resolveImage ? resolveImage(url) : ResolveUrl(url, options);

    private static string? ResolveUrl(string url, MarkdownLoadOptions? options) {
        if (url.Length == 0) return string.Empty;
        if (string.IsNullOrWhiteSpace(url)) return null;
        url = url.Trim();

        if (IsWindowsDriveLike(url)) {
            if (options?.DisallowFileUrls == true) return null;
            if (options?.RestrictUrlSchemes == true && !IsAllowedScheme("file", options.AllowedUrlSchemes)) return null;
            return url;
        }

        // Block scriptable schemes by default.
        if (TryGetScheme(url, out var scheme)) {
            if (options?.RestrictUrlSchemes == true && !IsAllowedScheme(scheme, options.AllowedUrlSchemes)) return null;
            if (options?.DisallowScriptUrls != false) {
                if (scheme.Equals("javascript", StringComparison.OrdinalIgnoreCase) ||
                    scheme.Equals("vbscript", StringComparison.OrdinalIgnoreCase)) {
                    return null;
                }
            }
            if (options?.DisallowFileUrls == true) {
                if (scheme.Equals("file", StringComparison.OrdinalIgnoreCase) || IsWindowsDriveLike(url)) return null;
            }
            if (scheme.Equals("mailto", StringComparison.OrdinalIgnoreCase)) return (options?.AllowMailtoUrls ?? true) ? url : null;
            if (scheme.Equals("data", StringComparison.OrdinalIgnoreCase)) return (options?.AllowDataUrls ?? true) ? url : null;
            // http/https and unknown schemes: keep as-is (host may further restrict)
            return url;
        }

        if (url.StartsWith("//", StringComparison.Ordinal)) {
            if (options?.AllowProtocolRelativeUrls != false) {
                if (options?.RestrictUrlSchemes == true && !IsAllowedScheme("http", options.AllowedUrlSchemes) && !IsAllowedScheme("https", options.AllowedUrlSchemes)) return null;
                return url;
            }
            return null;
        }
        if (url.StartsWith("#", StringComparison.Ordinal)) return url;

        var baseUri = options?.BaseUri;
        if (!string.IsNullOrWhiteSpace(baseUri)) {
            try {
                // Legacy behavior: only apply BaseUri when it is http(s), and only resolve into http(s).
                var baseAbs = new Uri(baseUri, UriKind.Absolute);
                if (!baseAbs.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !baseAbs.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url;
                }
                var resolved = new Uri(baseAbs, url);
                if (!resolved.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !resolved.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url;
                }
                return resolved.ToString();
            }
            catch (UriFormatException) { /* invalid base or relative path; keep original */ }
        }

        return url; // relative or unknown: leave as-is
    }

    private static bool IsAllowedScheme(string scheme, string[]? allowedSchemes) {
        if (string.IsNullOrEmpty(scheme)) return false;
        if (allowedSchemes == null || allowedSchemes.Length == 0) return false;
        for (int i = 0; i < allowedSchemes.Length; i++) {
            var s = allowedSchemes[i];
            if (string.IsNullOrWhiteSpace(s)) continue;
            if (scheme.Equals(s.Trim(), StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static bool TryGetScheme(string url, out string scheme) {
        scheme = string.Empty;
        int colon = url.IndexOf(':');
        if (colon <= 0) return false;
        // URI scheme must start with a letter and be [A-Za-z0-9+.-]*. None of those characters is a
        // path, query or fragment delimiter, so this also rejects a delimiter in front of the colon.
        char first = url[0];
        if (!char.IsLetter(first)) return false;
        for (int i = 1; i < colon; i++) {
            char c = url[i];
            bool ok = char.IsLetterOrDigit(c) || c == '+' || c == '-' || c == '.';
            if (!ok) return false;
        }
        scheme = GetSchemeText(url, colon);
        return true;
    }

    // The common schemes come back as shared strings rather than a substring per URL.
    private static string GetSchemeText(string url, int length) {
        string? known = length switch {
            4 => MatchScheme(url, "http") ?? MatchScheme(url, "data") ?? MatchScheme(url, "file"),
            5 => MatchScheme(url, "https"),
            6 => MatchScheme(url, "mailto"),
            _ => null
        };
        return known ?? url.Substring(0, length);
    }

    private static string? MatchScheme(string url, string scheme) =>
        string.CompareOrdinal(url, 0, scheme, 0, scheme.Length) == 0 ? scheme : null;

    private static bool IsWindowsDriveLike(string url) {
        // Treat "C:\..." and "C:/..." as file-like.
        return url.Length >= 3
               && char.IsLetter(url[0])
               && url[1] == ':'
               && (url[2] == '\\' || url[2] == '/');
    }
}
