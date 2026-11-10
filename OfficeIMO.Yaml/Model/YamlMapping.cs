using System.Collections;

namespace OfficeIMO.Yaml;

/// <summary>One key and value of a <see cref="YamlMapping"/>. Keys are nodes, so complex keys are representable.</summary>
public readonly struct YamlMappingEntry {
    internal YamlMappingEntry(YamlNode key, YamlNode value) {
        Key = key;
        Value = value;
    }

    /// <summary>Gets the key node.</summary>
    public YamlNode Key { get; }

    /// <summary>Gets the value node; an omitted value is an empty plain scalar.</summary>
    public YamlNode Value { get; }
}

/// <summary>
/// A mapping node: block (<c>key: value</c>) or flow (<c>{a: 1}</c>). Entries keep source order, including equal keys
/// a lenient <see cref="YamlDuplicateKeyPolicy"/> admitted.
/// </summary>
/// <remarks>
/// String-key lookups (the indexer, <see cref="TryGetValue"/>, <see cref="ContainsKey"/> and paths) are the mapping's
/// typed view. They match scalar keys, including keys reached through an alias, by decoded text, so <c>"a"</c> and
/// <c>a</c> are the same key, and they are O(1) after the first lookup builds an index. Under the YAML 1.1 schema they
/// also see the entries of <c>&lt;&lt;</c> merge keys: the mapping's own keys win, then merged mappings in order.
/// Enumeration and <see cref="Count"/> show the tree as written; merges are never flattened into it.
/// </remarks>
public sealed class YamlMapping : YamlNode, IReadOnlyList<YamlMappingEntry> {
    /// <summary>Mappings up to this size are searched linearly instead of through a dictionary.</summary>
    private const int LinearLookupLimit = 8;

    internal YamlMappingEntry[] _entries = Array.Empty<YamlMappingEntry>();
    internal int _count;
    internal int _contentStart;
    private YamlKeyIndex? _index;

    internal YamlMapping(YamlDocumentPart part, in YamlEvent ev) : base(part, ev.Style, ev.Start, ev.End) {
        _contentStart = ev.End;
    }

    /// <summary>Gets the number of entries as written, without merged entries.</summary>
    public int Count => _count;

    /// <summary>Gets the entry at a zero-based position in source order.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the mapping.</exception>
    public YamlMappingEntry this[int index] {
        get {
            if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the mapping.");
            return _entries[index];
        }
    }

    /// <summary>Gets the value of the key with the given text. Aliases are returned as <see cref="YamlAlias"/> nodes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="KeyNotFoundException">No key has the text, including merged keys.</exception>
    public YamlNode this[string key] {
        get {
            if (TryGetValue(key, out YamlNode? value)) return value!;
            throw new KeyNotFoundException("The mapping at " + Start + " has no key '" + key + "'.");
        }
    }

    /// <summary>Gets the value of the key with the given text, looking through merge keys under the YAML 1.1 schema.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    public bool TryGetValue(string key, out YamlNode? value) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        return TryGetValueCore(key, out value);
    }

    /// <summary>Returns whether a key has the given text, including merged keys under the YAML 1.1 schema.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    public bool ContainsKey(string key) => TryGetValue(key, out _);

    /// <summary>Returns an enumerator over the entries as written that does not allocate.</summary>
    public Enumerator GetEnumerator() => new Enumerator(this);

    IEnumerator<YamlMappingEntry> IEnumerable<YamlMappingEntry>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private bool TryGetValueCore(string key, out YamlNode? value) {
        YamlKeyIndex index = _index ??= YamlKeyIndex.Build(this);
        int found = index.Map is null ? FindLinear(key) : index.Map.TryGetValue(key, out int i) ? i : -1;
        if (found >= 0) {
            value = _entries[found].Value;
            return true;
        }
        int[] merges = index.Merges;
        for (int m = 0; m < merges.Length; m++) {
            YamlNode source = _entries[merges[m]].Value.Resolved;
            if (source is YamlMapping merged) {
                if (merged.TryGetValueCore(key, out value)) return true;
            } else if (source is YamlSequence list) {
                for (int s = 0; s < list._count; s++) {
                    if (list._items[s].Resolved is YamlMapping item && item.TryGetValueCore(key, out value)) return true;
                }
            }
        }
        value = null;
        return false;
    }

    private int FindLinear(string key) {
        int found = -1;
        bool last = _part._document.DuplicateKeys == YamlDuplicateKeyPolicy.LastWins;
        for (int i = 0; i < _count; i++) {
            if (string.Equals(KeyText(_entries[i].Key), key, StringComparison.Ordinal)) {
                if (!last) return i;
                found = i;
            }
        }
        return found;
    }

    /// <summary>Returns the decoded text a string lookup matches for a key, or null for collection keys.</summary>
    internal static string? KeyText(YamlNode key) => key.Resolved is YamlScalar scalar ? scalar.Value : null;

    internal int IndexOfNode(YamlNode node, out bool isKey) {
        for (int i = 0; i < _count; i++) {
            if (ReferenceEquals(_entries[i].Key, node)) { isKey = true; return i; }
            if (ReferenceEquals(_entries[i].Value, node)) { isKey = false; return i; }
        }
        isKey = false;
        return -1;
    }

    /// <summary>The lazily built lookup state of one mapping: a dictionary for large mappings and the merge entries.</summary>
    private sealed class YamlKeyIndex {
        private static readonly YamlKeyIndex Empty = new YamlKeyIndex(null, Array.Empty<int>());

        private YamlKeyIndex(Dictionary<string, int>? map, int[] merges) {
            Map = map;
            Merges = merges;
        }

        public Dictionary<string, int>? Map { get; }

        public int[] Merges { get; }

        public static YamlKeyIndex Build(YamlMapping mapping) {
            int count = mapping._count;
            YamlMappingEntry[] entries = mapping._entries;
            int merges = 0;
            for (int i = 0; i < count; i++) {
                if (entries[i].Key is YamlScalar key && key.IsMergeKey()) merges++;
            }
            Dictionary<string, int>? map = null;
            if (count > LinearLookupLimit) {
                bool last = mapping._part._document.DuplicateKeys == YamlDuplicateKeyPolicy.LastWins;
                map = new Dictionary<string, int>(count, StringComparer.Ordinal);
                for (int i = 0; i < count; i++) {
                    string? text = KeyText(entries[i].Key);
                    if (text is null) continue;
                    if (last) map[text] = i;
                    else if (!map.ContainsKey(text)) map.Add(text, i);
                }
            }
            if (map is null && merges == 0) return Empty;
            int[] mergeEntries = merges == 0 ? Array.Empty<int>() : new int[merges];
            for (int i = 0, m = 0; m < merges; i++) {
                if (entries[i].Key is YamlScalar key && key.IsMergeKey()) mergeEntries[m++] = i;
            }
            return new YamlKeyIndex(map, mergeEntries);
        }
    }

    /// <summary>Enumerates the entries of a <see cref="YamlMapping"/> as written.</summary>
    public struct Enumerator : IEnumerator<YamlMappingEntry> {
        private readonly YamlMapping _mapping;
        private int _index;

        internal Enumerator(YamlMapping mapping) {
            _mapping = mapping;
            _index = -1;
        }

        /// <inheritdoc />
        public YamlMappingEntry Current => _mapping._entries[_index];

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() => ++_index < _mapping._count;

        /// <inheritdoc />
        public void Reset() => _index = -1;

        /// <inheritdoc />
        public void Dispose() { }
    }
}
