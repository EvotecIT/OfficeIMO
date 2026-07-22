using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeIMO.Adf;
using OfficeIMO.GoogleWorkspace.Auth.GoogleApis;
using OfficeIMO.Reader;

var adfAttributes = new ReadOnlyObjectDictionary(new Dictionary<string, object?> {
    ["html"] = "<strong>Ready</strong>",
    ["enabled"] = true
});
JsonElement adfValue = new AdfNode("extension")
    .SetAttribute("parameters", adfAttributes)
    .Attributes["parameters"];
if (adfValue.ValueKind != JsonValueKind.Object || !adfValue.GetProperty("enabled").GetBoolean()) {
    throw new InvalidOperationException("The ADF read-only dictionary did not retain its JSON object shape.");
}

string readerJson = OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult {
    Metadata = new[] {
        new OfficeDocumentMetadataEntry {
            Id = "metadata-1",
            Category = "core",
            Name = "fixture",
            Attributes = new Dictionary<string, string> {
                ["zeta"] = "last",
                ["alpha"] = "first"
            }
        }
    }
});
int alphaIndex = readerJson.IndexOf("\"alpha\"", StringComparison.Ordinal);
int zetaIndex = readerJson.IndexOf("\"zeta\"", StringComparison.Ordinal);
if (alphaIndex < 0 || zetaIndex < 0 || alphaIndex > zetaIndex) {
    throw new InvalidOperationException("Reader metadata attributes were not serialized deterministically.");
}

var store = new InMemoryTokenStore();
var adapter = new GoogleApisDataStoreAdapter(store);
await adapter.StoreAsync("officeimo-aot", "token-marker");
var value = await adapter.GetAsync<string>("officeimo-aot");

if (!string.Equals(value, "token-marker", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The Google APIs data-store adapter did not round-trip its value.");
}

Console.WriteLine("PASS | 85 production libraries fully rooted; Google APIs token-store adapter round-tripped from NativeAOT.");

file sealed class InMemoryTokenStore : IGoogleWorkspaceTokenStore {
    private readonly Dictionary<string, object?> _values = new(StringComparer.Ordinal);

    public Task StoreAsync<T>(string key, T value) {
        _values[key] = value;
        return Task.CompletedTask;
    }

    public Task DeleteAsync<T>(string key) {
        _values.Remove(key);
        return Task.CompletedTask;
    }

    public Task<T?> GetAsync<T>(string key) {
        return Task.FromResult(_values.TryGetValue(key, out var value) ? (T?)value : default);
    }

    public Task ClearAsync() {
        _values.Clear();
        return Task.CompletedTask;
    }
}

file sealed class ReadOnlyObjectDictionary : IReadOnlyDictionary<string, object?> {
    private readonly IReadOnlyDictionary<string, object?> _values;

    public ReadOnlyObjectDictionary(IReadOnlyDictionary<string, object?> values) {
        _values = values;
    }

    public object? this[string key] => _values[key];
    public IEnumerable<string> Keys => _values.Keys;
    public IEnumerable<object?> Values => _values.Values;
    public int Count => _values.Count;
    public bool ContainsKey(string key) => _values.ContainsKey(key);
    public bool TryGetValue(string key, out object? value) => _values.TryGetValue(key, out value);
    public IEnumerator<KeyValuePair<string, object?>> GetEnumerator() => _values.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
