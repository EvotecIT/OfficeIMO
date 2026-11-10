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
/// <example>
/// <code>
/// YamlMapping spec = document.Select("spec")!.AsMapping();
/// spec["replicas"] = 3;                                // only the value's text changes
/// spec.Add("paused", false).SetComment("rollout hold");
/// spec.Remove("progressDeadlineSeconds");             // its comments and line go with it
/// </code>
/// </example>
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
    /// <summary>The loaded entries as (start, end) source pairs, kept from the first insertion or removal so saving can find their source text.</summary>
    internal int[]? _original;
    private YamlKeyIndex? _keyIndex;

    internal YamlMapping(YamlDocumentPart part, in YamlEvent ev) : base(part, ev.Style, ev.Start, ev.End) {
        _contentStart = ev.End;
    }

    /// <summary>Creates an empty block mapping. Inside a flow collection it is written as flow.</summary>
    public YamlMapping() : base(YamlStyle.Block) {
    }

    /// <summary>Creates an empty mapping with a style: <see cref="YamlStyle.Block"/> or <see cref="YamlStyle.Flow"/> (<c>{a: 1}</c>).</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="style"/> is not a collection style.</exception>
    public YamlMapping(YamlStyle style) : base(YamlEdit.CheckCollectionStyle(style)) {
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

    /// <summary>
    /// Gets the value of the key with the given text, or sets it as <see cref="Set"/> does. Aliases are returned as
    /// <see cref="YamlAlias"/> nodes.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="KeyNotFoundException">Getting: no key has the text, including merged keys.</exception>
    public YamlNode this[string key] {
        get {
            if (TryGetValue(key, out YamlNode? value)) return value!;
            throw new KeyNotFoundException("The mapping" + Where() + " has no key '" + key + "'.");
        }
        set => Set(key, value);
    }

    /// <summary>
    /// Sets the value of a key, adding the entry at the end when the mapping has no such key of its own. A key that only
    /// a <c>&lt;&lt;</c> merge provides gets an entry of its own, which overrides the merged one; merged mappings are never
    /// edited. A scalar set over a scalar updates it in place, keeping its anchor, tag, comments and, for an equal value,
    /// its text; other values replace the node.
    /// </summary>
    /// <param name="key">The key text.</param>
    /// <param name="value">The value; <see langword="null"/> sets a null scalar. Strings, numbers and booleans convert implicitly.</param>
    /// <returns>The node now at the key.</returns>
    /// <exception cref="InvalidOperationException"><paramref name="value"/> already belongs to a collection.</exception>
    public YamlNode Set(string key, YamlNode? value) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        int index = IndexOf(key);
        if (index < 0) return Insert(_count, key, value);
        return YamlEdit.ReplaceValue(this, index, value ?? new YamlScalar((string?)null));
    }

    /// <summary>Adds an entry at the end.</summary>
    /// <returns>The value node, now in the mapping.</returns>
    /// <exception cref="ArgumentException">The mapping already has the key.</exception>
    /// <exception cref="InvalidOperationException"><paramref name="value"/> already belongs to a collection.</exception>
    public YamlNode Add(string key, YamlNode? value) => Insert(_count, key, value);

    /// <summary>Adds an entry with a key node at the end, for keys with a chosen style, a tag, or a collection as key.</summary>
    /// <returns>The value node, now in the mapping.</returns>
    public YamlNode Add(YamlNode key, YamlNode? value) => Insert(_count, key, value);

    /// <summary>
    /// Inserts an entry before the entry at <paramref name="index"/>, or at the end when it equals <see cref="Count"/>.
    /// The new entry takes the indentation, layout and quoting of its neighbors.
    /// </summary>
    /// <returns>The value node, now in the mapping.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside 0..<see cref="Count"/>.</exception>
    /// <exception cref="ArgumentException">The mapping already has the key.</exception>
    public YamlNode Insert(int index, string key, YamlNode? value) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        return Insert(index, new YamlScalar(key), value);
    }

    /// <summary>Inserts an entry with a key node; see <see cref="Insert(int, string, YamlNode?)"/>.</summary>
    /// <returns>The value node, now in the mapping.</returns>
    public YamlNode Insert(int index, YamlNode key, YamlNode? value) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        if ((uint)index > (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the mapping.");
        return YamlEdit.InsertEntry(this, index, key, value ?? new YamlScalar((string?)null));
    }

    /// <summary>
    /// Removes the entry with the given key text, together with the comment lines above it, its line and its trailing
    /// comment. Neighbors keep their comments and blank lines. Merged keys are never removed from their source.
    /// </summary>
    /// <returns>Whether the mapping had the key.</returns>
    /// <exception cref="InvalidOperationException">Aliases elsewhere refer to an anchor inside the entry.</exception>
    public bool Remove(string key) {
        int index = IndexOf(key ?? throw new ArgumentNullException(nameof(key)));
        if (index < 0) return false;
        RemoveAt(index);
        return true;
    }

    /// <summary>Removes the entry at a position; see <see cref="Remove(string)"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the mapping.</exception>
    public void RemoveAt(int index) {
        if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the mapping.");
        YamlEdit.RemoveChild(this, index);
    }

    /// <summary>
    /// Renames a key, keeping its entry's position, value and comments. The key keeps its quoting when that can represent
    /// the new text.
    /// </summary>
    /// <returns>The key node.</returns>
    /// <exception cref="KeyNotFoundException">The mapping has no key <paramref name="key"/> of its own.</exception>
    /// <exception cref="ArgumentException">The mapping already has a key <paramref name="newKey"/>.</exception>
    public YamlScalar RenameKey(string key, string newKey) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        if (newKey is null) throw new ArgumentNullException(nameof(newKey));
        int index = IndexOf(key);
        if (index < 0) throw new KeyNotFoundException("The mapping" + Where() + " has no key '" + key + "' of its own.");
        if (_entries[index].Key.Resolved is not YamlScalar scalar || _entries[index].Key is YamlAlias) {
            throw new InvalidOperationException("Only a scalar key can be renamed.");
        }
        if (key == newKey) return scalar;
        if (IndexOf(newKey) >= 0) throw new ArgumentException("The mapping already has the key '" + newKey + "'.", nameof(newKey));
        scalar.SetValue(newKey);
        return scalar;
    }

    /// <summary>Returns the position of the mapping's own entry with the key text, or -1. Merged keys are not counted.</summary>
    public int IndexOf(string key) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        YamlKeyIndex index = _keyIndex ??= YamlKeyIndex.Build(this);
        return index.Map is null ? FindLinear(key) : index.Map.TryGetValue(key, out int i) ? i : -1;
    }

    /// <summary>Records each child's position from <paramref name="from"/> on, so lookups of a node's entry take constant time.</summary>
    internal void Renumber(int from) {
        for (int i = from; i < _count; i++) {
            _entries[i].Key._index = i;
            _entries[i].Value._index = i;
        }
    }

    /// <summary>Drops the key lookup index after keys change.</summary>
    internal void ResetIndex() => _keyIndex = null;

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
        YamlKeyIndex index = _keyIndex ??= YamlKeyIndex.Build(this);
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
        bool last = Document?.DuplicateKeys == YamlDuplicateKeyPolicy.LastWins;
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
        int known = node._index;
        if ((uint)known < (uint)_count) {
            if (ReferenceEquals(_entries[known].Key, node)) { isKey = true; return known; }
            if (ReferenceEquals(_entries[known].Value, node)) { isKey = false; return known; }
        }
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
                bool last = mapping.Document?.DuplicateKeys == YamlDuplicateKeyPolicy.LastWins;
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
