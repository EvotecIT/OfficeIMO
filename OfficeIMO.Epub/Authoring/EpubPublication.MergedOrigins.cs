namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private readonly Dictionary<string, HashSet<string>> _mergedEntryOrigins = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);

    private void RetainMergedOrigins(string destination, string source) {
        if (!_mergedEntryOrigins.TryGetValue(destination, out HashSet<string>? origins))
            _mergedEntryOrigins.Add(destination, origins = new HashSet<string>(StringComparer.Ordinal));
        if (_entryOrigins.TryGetValue(source, out string? original)) origins.Add(original);
        if (_mergedEntryOrigins.TryGetValue(source, out HashSet<string>? previous)) origins.UnionWith(previous);
        _entryOrigins.Remove(source); _mergedEntryOrigins.Remove(source);
    }
}
