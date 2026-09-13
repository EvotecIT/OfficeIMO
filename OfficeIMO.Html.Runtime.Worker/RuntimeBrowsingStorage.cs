namespace OfficeIMO.Html.Runtime.Worker;

// The two storage areas are session-local and partitioned by origin. Limits are
// cumulative across origins, so visiting another origin cannot reset a quota.
internal sealed class RuntimeBrowsingStorage(int maximumCharacters) {
    private readonly Dictionary<(string Origin, bool Local), Dictionary<string, string>> _areas = new();
    private readonly long[] _sizes = new long[2];
    private readonly object _sync = new();

    internal string[][] Read(string origin, bool local) {
        lock (_sync) return Area(origin, local).Select(pair => new[] { pair.Key, pair.Value }).ToArray();
    }

    internal void Write(string origin, bool local, string operation, string key, string value) {
        lock (_sync) {
            var area = Area(origin, local);
            int slot = local ? 0 : 1;
            if (operation == "clear") { _sizes[slot] -= area.Sum(pair => (long)pair.Key.Length + pair.Value.Length); area.Clear(); return; }
            area.TryGetValue(key, out var previous);
            if (operation == "remove") {
                if (previous != null) { area.Remove(key); _sizes[slot] -= key.Length + previous.Length; }
                return;
            }
            long size = _sizes[slot] + value.Length + (previous == null ? key.Length : -previous.Length);
            if (size > maximumCharacters) throw new HtmlScriptRuntimeException("The session storage area exceeds MaxStorageCharacters.");
            area[key] = value;
            _sizes[slot] = size;
        }
    }

    private Dictionary<string, string> Area(string origin, bool local) {
        if (!_areas.TryGetValue((origin, local), out var area)) _areas.Add((origin, local), area = new(StringComparer.Ordinal));
        return area;
    }
}
