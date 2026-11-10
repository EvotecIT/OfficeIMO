namespace OfficeIMO.Yaml;

/// <summary>
/// A small direct-mapped cache that returns the same string for repeated short spans, such as mapping keys and anchor
/// names. A collision replaces the slot, so memory stays bounded regardless of input. The slots are rented, so a
/// short read that meets one key does not allocate the whole table; <see cref="Release"/> clears and returns them.
/// </summary>
internal sealed class YamlStringCache {
    private const int Size = 256;
    private const int MaxLength = 64;
    private string?[] _entries = ArrayPool<string?>.Shared.Rent(Size);

    public string Get(ReadOnlySpan<char> text) {
        if (text.Length == 0) return string.Empty;
        if (text.Length > MaxLength) return text.ToString();
        uint hash = 2166136261;
        foreach (char ch in text) hash = (hash ^ ch) * 16777619;
        int slot = (int)((hash ^ (hash >> 16)) & (Size - 1));
        string? entry = _entries[slot];
        if (entry is not null && text.SequenceEqual(entry.AsSpan())) return entry;
        entry = text.ToString();
        _entries[slot] = entry;
        return entry;
    }

    public void Release() {
        string?[] entries = _entries;
        if (entries.Length == 0) return;
        _entries = Array.Empty<string?>();
        ArrayPool<string?>.Shared.Return(entries, clearArray: true);
    }
}
