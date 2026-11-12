using System.Runtime.InteropServices;

namespace OfficeIMO.Yaml;

/// <summary>
/// A small direct-mapped cache that returns the same string for repeated short spans, such as mapping keys and anchor
/// names. A collision replaces the slot, so memory stays bounded regardless of input. The slots are rented, so a
/// short read that meets one key does not allocate the whole table; <see cref="Release"/> clears the slots it filled,
/// tracked in a bit set. The last released cache stays with its thread for the next reader, so reading many small
/// documents neither allocates caches nor rents tables.
/// </summary>
internal sealed class YamlStringCache {
    private const int Size = 256;
    private const int MaxLength = 64;
    [ThreadStatic] private static YamlStringCache? t_spare;
    private string?[] _entries = ArrayPool<string?>.Shared.Rent(Size);
    private ulong _filled0, _filled1, _filled2, _filled3;

    private YamlStringCache() {
    }

    /// <summary>Takes this thread's released cache, or creates one.</summary>
    public static YamlStringCache Rent() {
        YamlStringCache? spare = t_spare;
        if (spare is null) return new YamlStringCache();
        t_spare = null;
        return spare;
    }

    public string Get(ReadOnlySpan<char> text) {
        if (text.Length == 0) return string.Empty;
        if (text.Length > MaxLength) return text.ToString();
        int slot = Slot(text);
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

    /// <summary>
    /// Hashes the span from its first and last eight bytes (overlapping for short text) and its length: two reads and two
    /// multiplications instead of a step per character. A collision only replaces a slot.
    /// </summary>
    private static int Slot(ReadOnlySpan<char> text) {
        ReadOnlySpan<byte> bytes = MemoryMarshal.AsBytes(text);
        ulong head, tail;
        if (bytes.Length >= 8) {
            head = MemoryMarshal.Read<ulong>(bytes);
            tail = MemoryMarshal.Read<ulong>(bytes.Slice(bytes.Length - 8));
        } else if (bytes.Length >= 4) {
            head = MemoryMarshal.Read<uint>(bytes);
            tail = MemoryMarshal.Read<uint>(bytes.Slice(bytes.Length - 4));
        } else {
            head = text[0];
            tail = 0;
        }
        ulong hash = (head * 0x9E3779B97F4A7C15UL) ^ (tail * 0xC2B2AE3D27D4EB4FUL) ^ (ulong)bytes.Length;
        return (int)((hash ^ (hash >> 29) ^ (hash >> 47)) & (Size - 1));
    }

    /// <summary>Clears the filled slots and keeps the cache for this thread's next reader, or returns its table. The caller drops its reference.</summary>
    public void Release() {
        string?[] entries = _entries;
        if (entries.Length == 0) return;
        Clear(entries, 0, _filled0);
        Clear(entries, 64, _filled1);
        Clear(entries, 128, _filled2);
        Clear(entries, 192, _filled3);
        _filled0 = _filled1 = _filled2 = _filled3 = 0;
        if (t_spare is null) {
            t_spare = this;
            return;
        }
        _entries = Array.Empty<string?>();
        ArrayPool<string?>.Shared.Return(entries);
    }

    private static void Clear(string?[] entries, int offset, ulong filled) {
        for (int i = 0; filled != 0; i++, filled >>= 1) {
            if ((filled & 1) != 0) entries[offset + i] = null;
        }
    }
}
