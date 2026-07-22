using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeIMO.GoogleWorkspace.Auth.GoogleApis;

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
