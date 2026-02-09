namespace OfficeIMO.Markdown;

internal static class HostAllowList {
    internal static bool IsAllowed(string? host, System.Collections.Generic.IReadOnlyList<string>? patterns) {
        if (patterns == null || patterns.Count == 0) return true;
        host ??= string.Empty;
        if (host.Trim().Length == 0) return false;

        var h = host.Trim().TrimEnd('.').ToLowerInvariant();
        if (h.Length == 0) return false;

        for (int i = 0; i < patterns.Count; i++) {
            var raw = patterns[i];
            if (raw == null) continue;
            var p = raw.Trim().TrimEnd('.').ToLowerInvariant();
            if (p.Length == 0) continue;

            if (p == "*") return true;

            if (p.StartsWith("*.", StringComparison.Ordinal)) {
                // Subdomains only.
                var suffix = p.Substring(1); // ".example.com"
                if (h.EndsWith(suffix, StringComparison.Ordinal) && h.Length > suffix.Length) return true;
                continue;
            }

            if (p.StartsWith(".", StringComparison.Ordinal)) {
                // Apex + any subdomain.
                var apex = p.Substring(1);
                if (h.Equals(apex, StringComparison.Ordinal)) return true;
                if (h.EndsWith(p, StringComparison.Ordinal) && h.Length > p.Length) return true;
                continue;
            }

            if (h.Equals(p, StringComparison.Ordinal)) return true;
        }

        return false;
    }
}
