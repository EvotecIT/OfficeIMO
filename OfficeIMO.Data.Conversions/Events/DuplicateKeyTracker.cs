namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Detects a repeated key within one object without allocating per key: names of every open object live in one
/// character arena, indexed by hash. Typical records have a few dozen keys, so a linear scan of hashes is cheapest.
/// </summary>
internal sealed class DuplicateKeyTracker {
    private char[] _arena = new char[1024];
    private int _arenaLength;
    private Entry[] _entries = new Entry[64];
    private int _entryCount;
    private int[] _frames = new int[16];
    private int[] _arenaFrames = new int[16];
    private int _depth;

    private struct Entry {
        public int Hash;
        public int Start;
        public int Length;
    }

    public void Push() {
        if (_depth == _frames.Length) {
            Array.Resize(ref _frames, _depth * 2);
            Array.Resize(ref _arenaFrames, _depth * 2);
        }

        _frames[_depth] = _entryCount;
        _arenaFrames[_depth] = _arenaLength;
        _depth++;
    }

    public void Pop() {
        if (_depth == 0) return;
        _depth--;
        _entryCount = _frames[_depth];
        _arenaLength = _arenaFrames[_depth];
    }

    public bool Contains(ReadOnlySpan<char> key) {
        if (_depth == 0) return false;
        int hash = Hash(key);
        for (int i = _frames[_depth - 1]; i < _entryCount; i++) {
            Entry entry = _entries[i];
            if (entry.Hash == hash && entry.Length == key.Length && key.SequenceEqual(new ReadOnlySpan<char>(_arena, entry.Start, entry.Length))) return true;
        }

        return false;
    }

    /// <summary>Adds a key to the innermost object; returns false when it is already there.</summary>
    public bool Add(ReadOnlySpan<char> key) {
        if (_depth == 0) return true;
        int hash = Hash(key);
        for (int i = _frames[_depth - 1]; i < _entryCount; i++) {
            Entry entry = _entries[i];
            if (entry.Hash == hash && entry.Length == key.Length && key.SequenceEqual(new ReadOnlySpan<char>(_arena, entry.Start, entry.Length))) {
                return false;
            }
        }

        if (_entryCount == _entries.Length) Array.Resize(ref _entries, _entries.Length * 2);
        if (_arena.Length - _arenaLength < key.Length) Array.Resize(ref _arena, Math.Max(_arena.Length * 2, _arenaLength + key.Length));
        key.CopyTo(new Span<char>(_arena, _arenaLength, key.Length));
        _entries[_entryCount++] = new Entry { Hash = hash, Start = _arenaLength, Length = key.Length };
        _arenaLength += key.Length;
        return true;
    }

    private static int Hash(ReadOnlySpan<char> key) {
        unchecked {
            int hash = (int)2166136261;
            foreach (char c in key) hash = (hash ^ c) * 16777619;
            return hash;
        }
    }
}
