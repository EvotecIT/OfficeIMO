namespace OfficeIMO.Yaml;

internal enum YamlRecordKind : byte {
    Scalar,
    Mapping,
    Sequence,
    Alias
}

/// <summary>
/// One loaded node. Records are stored in document order (pre-order), so a collection's first child is the next record
/// and the record after a subtree is <see cref="Size"/> records on. A mapping's children alternate key and value.
/// </summary>
internal struct YamlNodeRecord {
    /// <summary>Start of the node, at its first property (anchor or tag) when it has one.</summary>
    public int Start;
    public int End;
    /// <summary>
    /// Scalars: start of the raw value. Collections: where the content starts, after the properties and any opening
    /// bracket. Aliases: start of the anchor name after <c>*</c>.
    /// </summary>
    public int ValueStart;
    /// <summary>Scalars: end of the raw value. Aliases: end of the anchor name.</summary>
    public int ValueEnd;
    /// <summary>Records in the subtree, the node itself included.</summary>
    public int Size;
    /// <summary>Mappings: entries. Sequences: items. Aliases: the anchored record. Scalars: the block indentation.</summary>
    public int Count;
    /// <summary>The row of <see cref="YamlNodeTable.Properties"/> holding the anchor and tag, or -1.</summary>
    public int Props;
    public YamlRecordKind Kind;
    public YamlStyle Style;
    public YamlTokenFlags Flags;
    public YamlChomping Chomping;
}

/// <summary>Anchor, tag and alias facts of a loaded node; only nodes with an anchor or tag have a row.</summary>
internal struct YamlRecordProperties {
    /// <summary>Start of the anchor name after <c>&amp;</c> in the source, or -1; the name is read from the source when a node is created.</summary>
    public int AnchorStart;
    public int AnchorLength;
    public string? Tag;
    /// <summary>Nodes the anchored node stands for when every alias inside it is resolved, saturating.</summary>
    public long ExpandedSize;
    /// <summary>At least one alias refers to the node.</summary>
    public bool Aliased;
    /// <summary>The collection is still being loaded; an alias to it would make the graph cyclic.</summary>
    public bool Open;
}

/// <summary>
/// The loaded nodes of a <see cref="YamlDocument"/> as a compact table in pooled arrays. Loading writes one record per
/// node and creates no node objects; <see cref="YamlNode"/> objects are created when code first reaches them, one
/// collection's children at a time, and keep their identity from then on. Untouched subtrees therefore cost their
/// records only, and saving them copies their source text.
/// </summary>
/// <remarks>
/// Creating node objects is serialized by a lock so concurrent readers of a loaded document see one object per node.
/// The arrays go back to the pool when the document is disposed before any node object was created; once a node has
/// been handed out the table stays, because that node may still need its children.
/// </remarks>
internal sealed class YamlNodeTable {
    public YamlNodeRecord[] Records;
    public int Count;
    public YamlRecordProperties[] Properties;
    public int PropertyCount;
    private readonly string _source;
    /// <summary>Anchored nodes that aliases refer to, by record, once created; aliases find their targets here.</summary>
    private Dictionary<int, YamlNode>? _anchored;
    /// <summary>Shares the text of equal keys across the entries created from the table.</summary>
    private YamlStringCache? _keys;
    private bool _handedOut;
    private bool _released;

    public YamlNodeTable(string source) {
        _source = source;
        Records = ArrayPool<YamlNodeRecord>.Shared.Rent(Math.Max(16, source.Length / 12));
        Properties = ArrayPool<YamlRecordProperties>.Shared.Rent(8);
    }

    public int Add() {
        if (Count == Records.Length) Records = Grow(Records, Count, clear: false);
        return Count++;
    }

    public int AddProperties() {
        if (PropertyCount == Properties.Length) Properties = Grow(Properties, PropertyCount, clear: true);
        Properties[PropertyCount] = default;
        return PropertyCount++;
    }

    /// <summary>Returns the pooled arrays when no node object was created from them.</summary>
    public void Release() {
        lock (this) {
            if (_handedOut || _released) return;
            _released = true;
            Array.Clear(Properties, 0, PropertyCount);
            ArrayPool<YamlRecordProperties>.Shared.Return(Properties);
            ArrayPool<YamlNodeRecord>.Shared.Return(Records);
            Properties = Array.Empty<YamlRecordProperties>();
            Records = Array.Empty<YamlNodeRecord>();
            Count = PropertyCount = 0;
        }
    }

    private static T[] Grow<T>(T[] array, int count, bool clear) {
        T[] larger = ArrayPool<T>.Shared.Rent(count * 2);
        Array.Copy(array, larger, count);
        if (clear) Array.Clear(array, 0, count);
        ArrayPool<T>.Shared.Return(array);
        return larger;
    }

    public string? Tag(int record) => Records[record].Props < 0 ? null : Properties[Records[record].Props].Tag;

    /// <summary>The scalar's decoded value as a span: the raw source when it is the value, otherwise a decoded copy.</summary>
    public ReadOnlySpan<char> ValueSpan(int record) {
        ref YamlNodeRecord r = ref Records[record];
        if (YamlScalarDecoder.IsVerbatim(r.Style, r.Flags)) return _source.AsSpan(r.ValueStart, r.ValueEnd - r.ValueStart);
        return Value(record).AsSpan();
    }

    /// <summary>The scalar's decoded value.</summary>
    public string Value(int record) {
        ref YamlNodeRecord r = ref Records[record];
        return YamlScalar.DecodeValue(_source, r.ValueStart, r.ValueEnd, r.Style, r.Flags, r.Chomping, r.Count);
    }

#if !NETCOREAPP3_0_OR_GREATER
    private static readonly uint HashSeed = (uint)Guid.NewGuid().GetHashCode();
#endif

    /// <summary>
    /// Hashes text for the key indexes. On .NET Core the hash is the runtime's randomized string hash, so crafted keys
    /// cannot make lookups or duplicate checks collide on purpose; older frameworks use FNV-1a with a per-process seed.
    /// </summary>
    public static int HashText(ReadOnlySpan<char> text) {
#if NETCOREAPP3_0_OR_GREATER
        return string.GetHashCode(text);
#else
        uint hash = 2166136261 ^ HashSeed;
        foreach (char ch in text) hash = (hash ^ ch) * 16777619;
        return (int)(hash ^ (hash >> 15));
#endif
    }

    /// <summary>The record an alias refers to, or the record itself.</summary>
    public int Resolve(int record) => Records[record].Kind == YamlRecordKind.Alias ? Records[record].Count : record;

    /// <summary>The record after a node's subtree: its next sibling, when it has one.</summary>
    public int Next(int record) => record + Records[record].Size;

    /// <summary>The node object for the root of a part, created on first use.</summary>
    public YamlNode MaterializeRoot(YamlDocumentPart part, int record) {
        lock (this) {
            if (part._rootNode is { } existing) return existing;
            ThrowIfReleased();
            _handedOut = true;
            YamlNode root = Create(part, record, null, 0);
            Volatile.Write(ref part._rootNode, root);
            return root;
        }
    }

    /// <summary>
    /// Creates the entry array of a loaded mapping on first use, with each entry's key record in its slots. The entries'
    /// objects are created one entry at a time by <see cref="LoadEntry"/>, so reaching one key of a large mapping
    /// creates the objects of that entry only.
    /// </summary>
    public YamlMappingEntry[] CreateSlots(YamlMapping mapping) {
        lock (this) {
            if (mapping._entries is { } existing) return existing;
            ThrowIfReleased();
            var slots = new int[mapping._count];
            int child = mapping._record + 1;
            for (int i = 0; i < slots.Length; i++) {
                slots[i] = child;
                child = Next(Next(child));
            }
            mapping._slots = slots;
            var entries = new YamlMappingEntry[slots.Length];
            Volatile.Write(ref mapping._entries, entries);
            return entries;
        }
    }

    /// <summary>Creates the key and value objects of one entry of a loaded mapping.</summary>
    public YamlMappingEntry LoadEntry(YamlMapping mapping, int index) {
        lock (this) {
            YamlMappingEntry[] entries = mapping._entries!;
            YamlMappingEntry entry = entries[index];
            if (entry.Key is not null && entry.Value is not null) return entry;
            ThrowIfReleased();
            int key = mapping._slots![index];
            entry = new YamlMappingEntry(CreateKey(mapping, key, index), Create(mapping._part!, Next(key), mapping, index));
            entries[index] = entry;
            return entry;
        }
    }

    /// <summary>Creates the objects of every entry of a loaded mapping that has none yet.</summary>
    public void LoadEntries(YamlMapping mapping) {
        lock (this) {
            if (mapping._slots is null) return;
            for (int i = 0; i < mapping._count; i++) LoadEntry(mapping, i);
            Volatile.Write(ref mapping._slots, null);
        }
    }

    /// <summary>Creates the item array of a loaded sequence on first use, with each item's record in its slots.</summary>
    public YamlNode?[] CreateSlots(YamlSequence sequence) {
        lock (this) {
            if (sequence._items is { } existing) return existing;
            ThrowIfReleased();
            var slots = new int[sequence._count];
            int child = sequence._record + 1;
            for (int i = 0; i < slots.Length; i++) {
                slots[i] = child;
                child = Next(child);
            }
            sequence._slots = slots;
            var items = new YamlNode?[slots.Length];
            Volatile.Write(ref sequence._items, items);
            return items;
        }
    }

    /// <summary>Creates the object of one item of a loaded sequence.</summary>
    public YamlNode LoadItem(YamlSequence sequence, int index) {
        lock (this) {
            YamlNode?[] items = sequence._items!;
            if (items[index] is { } existing) return existing;
            ThrowIfReleased();
            YamlNode item = Create(sequence._part!, sequence._slots![index], sequence, index);
            Volatile.Write(ref items[index], item);
            return item;
        }
    }

    /// <summary>Creates the objects of every item of a loaded sequence that has none yet.</summary>
    public void LoadItems(YamlSequence sequence) {
        lock (this) {
            if (sequence._slots is null) return;
            for (int i = 0; i < sequence._count; i++) LoadItem(sequence, i);
            Volatile.Write(ref sequence._slots, null);
        }
    }

    /// <summary>Creates a key; a key's text is read now, shared with equal keys of other mappings, because keys are looked up by text.</summary>
    private YamlNode CreateKey(YamlMapping mapping, int record, int index) {
        YamlNode key = Create(mapping._part!, record, mapping, index);
        if (key is YamlScalar scalar && YamlScalarDecoder.IsVerbatim(scalar._style, scalar._flags)) {
            scalar._value = (_keys ??= YamlStringCache.Rent()).Get(_source.AsSpan(scalar._valueStart, scalar._valueEnd - scalar._valueStart));
        }
        return key;
    }

    /// <summary>
    /// Returns the decoded text of a key record, looking through an alias, or false for a collection key. The text is
    /// read from the source without a copy when the raw text is the value.
    /// </summary>
    public bool TryGetKeyText(int record, out ReadOnlySpan<char> text) {
        int resolved = Resolve(record);
        if (Records[resolved].Kind != YamlRecordKind.Scalar) {
            text = default;
            return false;
        }
        text = ValueSpan(resolved);
        return true;
    }

    /// <summary>Returns whether a scalar record is the merge key <c>&lt;&lt;</c> in a document where merge keys apply.</summary>
    public bool IsMergeKey(int record, bool mergeKeys) {
        ref YamlNodeRecord r = ref Records[record];
        if (r.Kind != YamlRecordKind.Scalar) return false;
        string? tag = Tag(record);
        if (tag is not null) return tag == YamlSchemaResolver.TagPrefix + "merge";
        if (r.Style != YamlStyle.Plain || !mergeKeys) return false;
        return r.ValueEnd - r.ValueStart == 2 && _source[r.ValueStart] == '<' && _source[r.ValueStart + 1] == '<';
    }

    /// <summary>
    /// Returns the object of a loaded node, creating the collections on the way to it. Aliases use it to reach their
    /// anchored node, which they find by record whether or not code has reached it yet.
    /// </summary>
    public YamlNode NodeAt(YamlDocumentPart part, int record) {
        lock (this) {
            if (_anchored is not null && _anchored.TryGetValue(record, out YamlNode? known)) return known;
            YamlNode current = part.RootNode!;
            while (current._record != record) {
                YamlNode? next = null;
                if (current is YamlMapping mapping) {
                    for (int i = 0; i < mapping._count && next is null; i++) {
                        if (!mapping.IsLoaded(i)) {
                            int key = mapping._slots![i];
                            if (record < key || record >= Next(Next(key))) continue;
                        }
                        YamlMappingEntry entry = mapping.EntryAt(i);
                        if (Contains(entry.Key, record)) next = entry.Key;
                        else if (Contains(entry.Value, record)) next = entry.Value;
                    }
                } else if (current is YamlSequence sequence) {
                    for (int i = 0; i < sequence._count && next is null; i++) {
                        if (!sequence.IsLoaded(i)) {
                            int item = sequence._slots![i];
                            if (record < item || record >= Next(item)) continue;
                        }
                        YamlNode node = sequence.ItemAt(i);
                        if (Contains(node, record)) next = node;
                    }
                }
                current = next ?? throw new InvalidOperationException("The anchored node of an alias is no longer in its document.");
            }
            return current;
        }
    }

    private bool Contains(YamlNode node, int record) => node._record >= 0 && node._record <= record && record < node._record + Records[node._record].Size;

    private YamlNode Create(YamlDocumentPart part, int record, YamlNode? parent, int position) {
        ref YamlNodeRecord r = ref Records[record];
        YamlNode node;
        switch (r.Kind) {
            case YamlRecordKind.Scalar:
                node = new YamlScalar(part, r.Style, r.Start, r.End, r.ValueStart, r.ValueEnd, r.Count, r.Chomping, r.Flags);
                break;
            case YamlRecordKind.Mapping:
                node = new YamlMapping(part, r.Style, r.Start, r.End, r.ValueStart, r.Count);
                break;
            case YamlRecordKind.Sequence:
                node = new YamlSequence(part, r.Style, r.Start, r.End, r.ValueStart, r.Count);
                break;
            default:
                node = new YamlAlias(part, r.Start, r.End, _source.Substring(r.ValueStart, r.ValueEnd - r.ValueStart), r.Count);
                break;
        }
        node._record = record;
        node._parent = parent;
        node._index = position;
        if (r.Props >= 0) {
            ref YamlRecordProperties p = ref Properties[r.Props];
            node._properties = new YamlNodeProperties { Anchor = p.AnchorStart < 0 ? null : _source.Substring(p.AnchorStart, p.AnchorLength), Tag = p.Tag };
            if (p.Aliased) {
                node._state |= YamlNodeState.Aliased;
                (_anchored ??= new Dictionary<int, YamlNode>())[record] = node;
            }
        }
        return node;
    }

    private void ThrowIfReleased() {
        if (_released) throw new ObjectDisposedException(nameof(YamlDocument), "The document was disposed before its nodes were read.");
    }
}
