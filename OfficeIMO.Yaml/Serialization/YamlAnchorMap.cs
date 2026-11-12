namespace OfficeIMO.Yaml;

/// <summary>
/// The anchors of the document being read, by name. An alias finds its anchor by the name's span in the source, so
/// looking it up creates no string; a redefined anchor replaces the earlier node for the aliases after it.
/// </summary>
internal sealed class YamlAnchorMap {
    private struct Entry {
        public string Name;
        public uint Hash;
        public int Next;
        public YamlRecording Value;
    }

    // Bucket heads are entry indexes plus one, so a cleared bucket array reads as empty.
    private int[] _buckets = new int[16];
    private Entry[] _entries = new Entry[8];
    private int _count;

    public int Count => _count;

    public void Set(string name, YamlRecording value) {
        uint hash = Hash(name.AsSpan());
        int bucket = (int)(hash & (uint)(_buckets.Length - 1));
        for (int i = _buckets[bucket] - 1; i >= 0; i = _entries[i].Next) {
            ref Entry entry = ref _entries[i];
            if (entry.Hash == hash && string.Equals(entry.Name, name, StringComparison.Ordinal)) {
                entry.Value = value;
                return;
            }
        }
        if (_count == _entries.Length) {
            Array.Resize(ref _entries, _count * 2);
            if (_count * 2 > _buckets.Length) {
                Rehash(_buckets.Length * 4);
                bucket = (int)(hash & (uint)(_buckets.Length - 1));
            }
        }
        _entries[_count] = new Entry { Name = name, Hash = hash, Next = _buckets[bucket] - 1, Value = value };
        _buckets[bucket] = ++_count;
    }

    public YamlRecording? Find(ReadOnlySpan<char> name) {
        if (_count == 0) return null;
        uint hash = Hash(name);
        for (int i = _buckets[(int)(hash & (uint)(_buckets.Length - 1))] - 1; i >= 0; i = _entries[i].Next) {
            ref Entry entry = ref _entries[i];
            if (entry.Hash == hash && name.SequenceEqual(entry.Name.AsSpan())) return entry.Value;
        }
        return null;
    }

    public void Clear() {
        if (_count == 0) return;
        Array.Clear(_buckets, 0, _buckets.Length);
        Array.Clear(_entries, 0, _count);
        _count = 0;
    }

    private void Rehash(int size) {
        _buckets = new int[size];
        for (int i = 0; i < _count; i++) {
            int bucket = (int)(_entries[i].Hash & (uint)(size - 1));
            _entries[i].Next = _buckets[bucket] - 1;
            _buckets[bucket] = i + 1;
        }
    }

    private static uint Hash(ReadOnlySpan<char> name) {
        uint hash = 2166136261;
        foreach (char ch in name) hash = (hash ^ ch) * 16777619;
        return hash ^ (hash >> 15);
    }
}
