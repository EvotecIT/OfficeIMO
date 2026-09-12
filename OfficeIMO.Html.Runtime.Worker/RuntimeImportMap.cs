using System.Text.Json;

namespace OfficeIMO.Html.Runtime.Worker;

// Import maps are parsed as data, with URL normalization and longest-prefix matching.
internal sealed class RuntimeImportMap {
    private readonly Dictionary<string, string?> _imports = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Dictionary<string, string?>> _scopes = new(StringComparer.Ordinal);
    private bool _loaded;
    private bool _resolving;
    internal void Seal() => _resolving = true;

    internal void Load(string source, Uri baseUrl) {
        if (_loaded || _resolving) throw new HtmlScriptRuntimeException("This module profile requires one import map before any module import.");
        using var json = JsonDocument.Parse(source);
        if (json.RootElement.ValueKind != JsonValueKind.Object) throw new HtmlScriptRuntimeException("An import map must be an object.");
        if (json.RootElement.TryGetProperty("imports", out var imports)) ReadEntries(imports, baseUrl, _imports);
        if (json.RootElement.TryGetProperty("scopes", out var scopes)) {
            if (scopes.ValueKind != JsonValueKind.Object) throw new HtmlScriptRuntimeException("Import map scopes must be an object.");
            foreach (var scope in scopes.EnumerateObject()) {
                var entries = new Dictionary<string, string?>(StringComparer.Ordinal);
                ReadEntries(scope.Value, baseUrl, entries);
                _scopes[new Uri(baseUrl, scope.Name).AbsoluteUri] = entries;
            }
        }
        _loaded = true;
    }

    internal Uri Resolve(string specifier, Uri referrer) {
        _resolving = true;
        string normalized = Normalize(specifier, referrer);
        foreach (var scope in _scopes.OrderByDescending(entry => entry.Key.Length)) {
            if ((referrer.AbsoluteUri == scope.Key || (scope.Key.EndsWith('/') && referrer.AbsoluteUri.StartsWith(scope.Key, StringComparison.Ordinal)))
                && Match(normalized, scope.Value, out var scoped)) return scoped!;
        }
        if (Match(normalized, _imports, out var mapped)) return mapped!;
        if (!UrlLike(specifier)) throw new HtmlScriptRuntimeException("The bare module specifier is not mapped: " + specifier);
        return new Uri(normalized, UriKind.Absolute);
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
}
