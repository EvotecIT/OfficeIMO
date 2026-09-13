using System.Text.Json;

namespace OfficeIMO.Html.Runtime.Worker;

// Import maps remain owned data. Resolution records let later maps add new rules
// without changing any specifier that the document has already resolved.
internal sealed class RuntimeImportMap {
    private readonly Dictionary<string, string?> _imports = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Dictionary<string, string?>> _scopes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _integrity = new(StringComparer.Ordinal);
    private readonly List<ResolutionRecord> _resolved = [];

    internal void Load(string source, Uri baseUrl) {
        using var json = JsonDocument.Parse(source);
        if (json.RootElement.ValueKind != JsonValueKind.Object) throw new HtmlScriptRuntimeException("An import map must be an object.");

        var imports = new Dictionary<string, string?>(StringComparer.Ordinal);
        var scopes = new Dictionary<string, Dictionary<string, string?>>(StringComparer.Ordinal);
        var integrity = new Dictionary<string, string>(StringComparer.Ordinal);
        if (json.RootElement.TryGetProperty("imports", out var importsJson)) ReadEntries(importsJson, baseUrl, imports);
        if (json.RootElement.TryGetProperty("scopes", out var scopesJson)) ReadScopes(scopesJson, baseUrl, scopes);
        if (json.RootElement.TryGetProperty("integrity", out var integrityJson)) ReadIntegrity(integrityJson, baseUrl, integrity);

        foreach (var scope in scopes) {
            if (!_scopes.TryGetValue(scope.Key, out var current)) _scopes.Add(scope.Key, current = new(StringComparer.Ordinal));
            foreach (var entry in scope.Value) {
                if (!current.ContainsKey(entry.Key) && !WouldChangeResolved(scope.Key, entry.Key)) current.Add(entry.Key, entry.Value);
            }
        }
        foreach (var entry in integrity) _integrity.TryAdd(entry.Key, entry.Value);
        foreach (var entry in imports) {
            if (!_imports.ContainsKey(entry.Key) && !WouldChangeResolved(entry.Key))
                _imports.Add(entry.Key, entry.Value);
        }
    }

    internal Uri Resolve(string specifier, Uri referrer) {
        string normalized = Normalize(specifier, referrer);
        Uri result;
        foreach (var scope in _scopes.OrderByDescending(entry => entry.Key.Length)) {
            if (ScopeMatches(scope.Key, referrer.AbsoluteUri) && Match(normalized, scope.Value, out var scoped)) {
                result = scoped!;
                Record(referrer, normalized);
                return result;
            }
        }
        if (Match(normalized, _imports, out var mapped)) result = mapped!;
        else if (!UrlLike(specifier)) throw new HtmlScriptRuntimeException("The bare module specifier is not mapped: " + specifier);
        else result = new Uri(normalized, UriKind.Absolute);
        Record(referrer, normalized);
        return result;
    }

    internal string? IntegrityFor(Uri url) => _integrity.GetValueOrDefault(url.AbsoluteUri);

    private void Record(Uri referrer, string specifier) {
        if (!_resolved.Any(record => record.Referrer == referrer.AbsoluteUri && record.Specifier == specifier))
            _resolved.Add(new(referrer.AbsoluteUri, specifier));
    }

    private bool WouldChangeResolved(string scope, string key) => _resolved.Any(record =>
        ScopeMatches(scope, record.Referrer) && WouldChange(record, key));

    private bool WouldChangeResolved(string key) => _resolved.Any(record => WouldChange(record, key));

    private static bool WouldChange(ResolutionRecord record, string key) =>
        key == record.Specifier || key.EndsWith('/') && record.Specifier.StartsWith(key, StringComparison.Ordinal);

    private static bool ScopeMatches(string scope, string referrer) =>
        referrer == scope || scope.EndsWith('/') && referrer.StartsWith(scope, StringComparison.Ordinal);

    private static void ReadScopes(JsonElement json, Uri baseUrl, Dictionary<string, Dictionary<string, string?>> scopes) {
        if (json.ValueKind != JsonValueKind.Object) throw new HtmlScriptRuntimeException("Import map scopes must be an object.");
        foreach (var scope in json.EnumerateObject()) {
            if (!Uri.TryCreate(baseUrl, scope.Name, out var scopeUrl)) continue;
            var entries = new Dictionary<string, string?>(StringComparer.Ordinal);
            ReadEntries(scope.Value, baseUrl, entries);
            scopes[scopeUrl.AbsoluteUri] = entries;
        }
    }

    private static void ReadIntegrity(JsonElement json, Uri baseUrl, Dictionary<string, string> entries) {
        if (json.ValueKind != JsonValueKind.Object) throw new HtmlScriptRuntimeException("Import map integrity must be an object.");
        foreach (var entry in json.EnumerateObject()) {
            if (UrlLike(entry.Name) && entry.Value.ValueKind == JsonValueKind.String)
                entries[new Uri(baseUrl, entry.Name).AbsoluteUri] = entry.Value.GetString()!;
        }
    }

    private static void ReadEntries(JsonElement json, Uri baseUrl, Dictionary<string, string?> entries) {
        if (json.ValueKind != JsonValueKind.Object) throw new HtmlScriptRuntimeException("Import map entries must be an object.");
        foreach (var entry in json.EnumerateObject()) {
            if (entry.Name.Length == 0) continue;
            string key = Normalize(entry.Name, baseUrl);
            string? target = null;
            if (entry.Value.ValueKind == JsonValueKind.String && UrlLike(entry.Value.GetString()!)) {
                target = new Uri(baseUrl, entry.Value.GetString()!).AbsoluteUri;
                if (key.EndsWith('/') && !target.EndsWith('/')) target = null;
            }
            entries[key] = target;
        }
    }

    private static bool Match(string specifier, Dictionary<string, string?> entries, out Uri? result) {
        foreach (var entry in entries.OrderByDescending(entry => entry.Key.Length)) {
            bool exact = specifier == entry.Key;
            if (!exact && !(entry.Key.EndsWith('/') && specifier.StartsWith(entry.Key, StringComparison.Ordinal))) continue;
            if (entry.Value == null) throw new HtmlScriptRuntimeException("The import map blocks module specifier: " + specifier);
            result = exact ? new Uri(entry.Value) : new Uri(new Uri(entry.Value), specifier[entry.Key.Length..]);
            if (!exact && !result.AbsoluteUri.StartsWith(entry.Value, StringComparison.Ordinal))
                throw new HtmlScriptRuntimeException("A module specifier escapes its import-map prefix.");
            return true;
        }
        result = null;
        return false;
    }

    private static bool UrlLike(string specifier) => specifier.StartsWith('/') || specifier.StartsWith("./", StringComparison.Ordinal)
        || specifier.StartsWith("../", StringComparison.Ordinal) || Uri.TryCreate(specifier, UriKind.Absolute, out _);
    private static string Normalize(string specifier, Uri baseUrl) => UrlLike(specifier) ? new Uri(baseUrl, specifier).AbsoluteUri : specifier;

    private sealed record ResolutionRecord(string Referrer, string Specifier);
}
