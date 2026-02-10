namespace OfficeIMO.Markdown;

internal static class HostAllowList {
    private static readonly System.Globalization.IdnMapping Idn = new System.Globalization.IdnMapping();

    private static string NormalizeHost(string? host) {
        host ??= string.Empty;
        var h = host.Trim().TrimEnd('.');
        if (h.Length == 0) return string.Empty;

        // Be forgiving for bracketed IPv6 input ("[::1]").
        if (h.Length >= 2 && h[0] == '[' && h[h.Length - 1] == ']') {
            h = h.Substring(1, h.Length - 2);
        }

        // Best-effort IDN normalization. For typical ASCII hosts, this is a no-op.
        if (NeedsIdn(h)) {
            try { h = Idn.GetAscii(h); } catch { /* best-effort */ }
        }

        return h.ToLowerInvariant();
    }

    private static bool NeedsIdn(string s) {
        for (int i = 0; i < s.Length; i++) {
            if (s[i] > 127) return true;
        }
        return false;
    }

    private static string NormalizePattern(string raw) {
        var p = (raw ?? string.Empty).Trim().TrimEnd('.');
        if (p.Length == 0) return string.Empty;
        if (p == "*") return "*";

        // Allow common misconfigurations: "https://example.com" -> "example.com".
        // Wildcard patterns won't parse as URIs, so this is best-effort.
        if (p.IndexOf("://", StringComparison.Ordinal) >= 0) {
            if (Uri.TryCreate(p, UriKind.Absolute, out var u) && u != null && !string.IsNullOrEmpty(u.Host)) {
                p = u.Host;
            }
        }

        // Allow "example.com:443" patterns; ignore port (Uri.Host never includes it).
        // Preserve IPv6 patterns.
        if (p.Length >= 2 && p[0] == '[') {
            // "[::1]:123" or "[::1]"
            int close = p.IndexOf(']');
            if (close > 0) {
                p = p.Substring(0, close + 1);
            }
        } else {
            int firstColon = p.IndexOf(':');
            if (firstColon >= 0 && p.IndexOf(':', firstColon + 1) < 0) {
                // Single-colon: treat as host:port.
                p = p.Substring(0, firstColon);
            }
        }

        // Preserve wildcard prefixes while normalizing the domain part.
        if (p.StartsWith("*.", StringComparison.Ordinal)) {
            var suffix = NormalizeHost(p.Substring(2));
            return "*." + suffix;
        }
        if (p.StartsWith(".", StringComparison.Ordinal)) {
            var apex = NormalizeHost(p.Substring(1));
            return "." + apex;
        }

        return NormalizeHost(p);
    }

    internal static bool IsAllowed(string? host, System.Collections.Generic.IReadOnlyList<string>? patterns) {
        if (patterns == null || patterns.Count == 0) return true;
        var h = NormalizeHost(host);
        if (h.Length == 0) return false;

        for (int i = 0; i < patterns.Count; i++) {
            var raw = patterns[i];
            if (raw == null) continue;
            var p = NormalizePattern(raw);
            if (p.Length == 0) continue;

            if (p == "*") return true;

            if (p.StartsWith("*.", StringComparison.Ordinal)) {
                // Subdomains only.
                var suffix = p.Substring(1); // ".example.com"
                if (h.EndsWith(suffix, StringComparison.Ordinal) && h.Length > suffix.Length) {
                    // Ensure we don't match odd strings if input isn't a real host; suffix begins with '.'.
                    int start = h.Length - suffix.Length;
                    if (start >= 0 && h[start] == '.') return true;
                }
                continue;
            }

            if (p.StartsWith(".", StringComparison.Ordinal)) {
                // Apex + any subdomain.
                var apex = p.Substring(1);
                if (h.Equals(apex, StringComparison.Ordinal)) return true;
                if (h.EndsWith(p, StringComparison.Ordinal) && h.Length > p.Length) {
                    int start = h.Length - p.Length;
                    if (start >= 0 && h[start] == '.') return true;
                }
                continue;
            }

            if (h.Equals(p, StringComparison.Ordinal)) return true;
        }

        return false;
    }
}
