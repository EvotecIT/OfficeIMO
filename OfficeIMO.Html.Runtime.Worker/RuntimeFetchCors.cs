namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeFetchCors {
    private static readonly HashSet<string> ExposedByDefault = new(StringComparer.OrdinalIgnoreCase) {
        "cache-control", "content-language", "content-length", "content-type", "expires", "last-modified", "pragma"
    };

    internal static void Check(IReadOnlyDictionary<string, string> headers, string origin) {
        if (!headers.TryGetValue("Access-Control-Allow-Origin", out string? allowed) || (allowed != "*" && allowed != origin))
            throw new HtmlScriptRuntimeException("The fetch response did not grant CORS access to the document origin.");
    }

    internal static void CheckPreflight(HtmlRuntimeResource response, string origin, string method, string[] names) {
        Check(response.Headers, origin);
        if (response.StatusCode < 200 || response.StatusCode >= 300) throw new HtmlScriptRuntimeException("CORS preflight failed.");
        var methods = Tokens(response.Headers, "Access-Control-Allow-Methods", StringComparer.Ordinal);
        if (method is not ("GET" or "HEAD" or "POST") && !methods.Contains(method) && !methods.Contains("*")) throw new HtmlScriptRuntimeException("CORS preflight did not allow the request method.");
        var allowed = Tokens(response.Headers, "Access-Control-Allow-Headers", StringComparer.OrdinalIgnoreCase);
        if (names.Any(name => !allowed.Contains(name) && (!allowed.Contains("*") || name == "authorization"))) throw new HtmlScriptRuntimeException("CORS preflight did not allow the request headers.");
    }

    internal static Dictionary<string, string> Expose(HtmlRuntimeResource response, bool crossOrigin) {
        var exposed = Tokens(response.Headers, "Access-Control-Expose-Headers", StringComparer.OrdinalIgnoreCase);
        return response.Headers.Where(h => !h.Key.Equals("set-cookie", StringComparison.OrdinalIgnoreCase) && !h.Key.Equals("set-cookie2", StringComparison.OrdinalIgnoreCase) &&
            (!crossOrigin || ExposedByDefault.Contains(h.Key) || exposed.Contains(h.Key) || exposed.Contains("*")))
            .ToDictionary(h => h.Key.ToLowerInvariant(), h => h.Value, StringComparer.Ordinal);
    }

    internal static string[] UnsafeHeaders(IReadOnlyDictionary<string, string> headers) {
        var unsafeNames = headers.Where(h => !Safe(h.Key, h.Value)).Select(h => h.Key.ToLowerInvariant()).ToHashSet(StringComparer.Ordinal);
        if (headers.Where(h => !unsafeNames.Contains(h.Key.ToLowerInvariant())).Sum(h => h.Value.Length) > 1024)
            unsafeNames.UnionWith(headers.Keys.Select(h => h.ToLowerInvariant()));
        return unsafeNames.OrderBy(h => h, StringComparer.Ordinal).ToArray();
    }

    private static bool Safe(string name, string value) {
        if (value.Length > 128) return false;
        switch (name.ToLowerInvariant()) {
            case "accept": return !UnsafeValue(value);
            case "accept-language":
            case "content-language": return value.All(c => char.IsAsciiLetterOrDigit(c) || " *,-.;=".Contains(c));
            case "content-type":
                if (UnsafeValue(value)) return false;
                return value.Split(';')[0].Trim().ToLowerInvariant() is "application/x-www-form-urlencoded" or "multipart/form-data" or "text/plain";
            case "range":
                if (!value.StartsWith("bytes=", StringComparison.Ordinal)) return false;
                string[] range = value[6..].Split('-');
                return range.Length == 2 && ulong.TryParse(range[0], System.Globalization.NumberStyles.None, null, out ulong start) &&
                    (range[1].Length == 0 || ulong.TryParse(range[1], System.Globalization.NumberStyles.None, null, out ulong end) && end >= start);
            default: return false;
        }
    }
    private static bool UnsafeValue(string value) => value.Any(c => c < 32 && c != '\t' || c == 127 || "\"():<>?@[\\]{}".Contains(c));
    private static HashSet<string> Tokens(IReadOnlyDictionary<string, string> headers, string name, StringComparer comparer) =>
        new(headers.TryGetValue(name, out string? value) ? value.Split(',').Select(item => item.Trim()).Where(item => item.Length != 0) : Array.Empty<string>(), comparer);
}
