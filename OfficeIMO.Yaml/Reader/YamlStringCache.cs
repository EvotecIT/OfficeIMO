namespace OfficeIMO.Yaml;

/// <summary>
/// A small direct-mapped cache that returns the same string for repeated short spans, such as mapping keys and anchor
/// names. A collision replaces the slot, so memory stays bounded regardless of input. The slots are rented, so a
/// short read that meets one key does not allocate the whole table; <see cref="Release"/> clears the slots it filled,
/// tracked in a bit set, and returns them.
/// </summary>
internal sealed class YamlStringCache {
    private const int Size = 256;
    private const int MaxLength = 64;
    private string?[] _entries = ArrayPool<string?>.Shared.Rent(Size);
    private ulong _filled0, _filled1, _filled2, _filled3;

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
        ulong bit = 1UL << (slot & 63);
        switch (slot >> 6) {
            case 0: _filled0 |= bit; break;
            case 1: _filled1 |= bit; break;
            case 2: _filled2 |= bit; break;
            default: _filled3 |= bit; break;
        }
        return entry;
    }

    public void Release() {
        string?[] entries = _entries;
        if (entries.Length == 0) return;
        _entries = Array.Empty<string?>();
        Clear(entries, 0, _filled0);
        Clear(entries, 64, _filled1);
        Clear(entries, 128, _filled2);
        Clear(entries, 192, _filled3);
        ArrayPool<string?>.Shared.Return(entries);
    }

    private static void Clear(string?[] entries, int offset, ulong filled) {
        for (int i = 0; filled != 0; i++, filled >>= 1) {
            if ((filled & 1) != 0) entries[offset + i] = null;
        }
    }
}
