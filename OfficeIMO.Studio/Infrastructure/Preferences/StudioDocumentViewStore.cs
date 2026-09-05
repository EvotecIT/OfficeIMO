using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>Retains at most 64 document presentation records, keyed by normalized path hash.</summary>
internal sealed class StudioDocumentViewStore {
    private const int MaximumEntries = 64;
    private const long MaximumFileBytes = 256 * 1024;
    private readonly string _path;
    private readonly Func<bool> _enabled;
    private List<Entry> _entries;

    internal StudioDocumentViewStore(string path, Func<bool>? enabled = null) {
        _path = Path.GetFullPath(path);
        _enabled = enabled ?? (() => true);
        _entries = Load();
    }

    internal StudioDocumentViewState Get(string documentPath) =>
        _enabled() ? _entries.FirstOrDefault(entry => entry.Key == Key(documentPath))?.State.Normalize() ?? new() : new();

    internal void Put(string documentPath, StudioDocumentViewState state) {
        if (!_enabled()) return;
        string key = Key(documentPath);
        var next = _entries.Where(entry => entry.Key != key).Take(MaximumEntries - 1).ToList();
        next.Insert(0, new Entry(key, state.Normalize()));
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        string temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            File.WriteAllBytes(temporary, JsonSerializer.SerializeToUtf8Bytes(new Snapshot(1, next)));
            File.Move(temporary, _path, overwrite: true);
            _entries = next;
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    internal void Clear() {
        File.Delete(_path);
        _entries = [];
    }

    private List<Entry> Load() {
        try {
            if (!File.Exists(_path) || new FileInfo(_path).Length > MaximumFileBytes) return [];
            Snapshot? snapshot = JsonSerializer.Deserialize<Snapshot>(File.ReadAllBytes(_path));
            if (snapshot?.SchemaVersion != 1 || snapshot.Entries is null) return [];
            return snapshot.Entries.Where(entry => entry is not null && entry.Key is { Length: 64 } &&
                    entry.Key.All(Uri.IsHexDigit) && entry.State is not null)
                .DistinctBy(entry => entry.Key).Take(MaximumEntries)
                .Select(entry => entry with { State = entry.State.Normalize() }).ToList();
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException) {
            return [];
        }
    }

    private static string Key(string path) {
        string normalized = Path.GetFullPath(path);
        if (OperatingSystem.IsWindows()) normalized = normalized.ToUpperInvariant();
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(normalized)));
    }

    private sealed record Entry(string Key, StudioDocumentViewState State);
    private sealed record Snapshot(int SchemaVersion, List<Entry> Entries);
}
