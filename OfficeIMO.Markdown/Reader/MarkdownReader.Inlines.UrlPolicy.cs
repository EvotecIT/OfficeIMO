namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static string? ResolveUrl(string url, MarkdownReaderOptions? options) {
        if (url.Length == 0) return string.Empty;
        if (string.IsNullOrWhiteSpace(url)) return null;
        url = url.Trim();

        if (IsWindowsDriveLike(url)) {
            return options?.DisallowFileUrls == true ? null : url;
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

        if (url.StartsWith("//")) {
            if (options?.AllowProtocolRelativeUrls != false) {
                if (options?.RestrictUrlSchemes == true && !IsAllowedScheme("http", options.AllowedUrlSchemes) && !IsAllowedScheme("https", options.AllowedUrlSchemes)) return null;
                return url;
            }
            return null;
        }
        if (url.StartsWith("#")) return url;

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
        // If there's a path/query/fragment delimiter before ':', it's not a scheme.
        int slash = url.IndexOfAny(new[] { '/', '?', '#' });
        if (slash >= 0 && slash < colon) return false;
        // URI scheme must start with a letter and be [A-Za-z0-9+.-]*
        char first = url[0];
        if (!char.IsLetter(first)) return false;
        for (int i = 1; i < colon; i++) {
            char c = url[i];
            bool ok = char.IsLetterOrDigit(c) || c == '+' || c == '-' || c == '.';
            if (!ok) return false;
        }
        scheme = url.Substring(0, colon);
        return true;
    }

    private static bool IsWindowsDriveLike(string url) {
        // Treat "C:\..." and "C:/..." as file-like.
        return url.Length >= 3
               && char.IsLetter(url[0])
               && url[1] == ':'
               && (url[2] == '\\' || url[2] == '/');
    }
}
