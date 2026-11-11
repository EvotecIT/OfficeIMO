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
    /// <summary>Scalars: start of the raw value. Collections: where the content starts, after the properties and any opening bracket.</summary>
    public int ValueStart;
    /// <summary>Scalars: end of the raw value.</summary>
    public int ValueEnd;
    /// <summary>Records in the subtree, the node itself included.</summary>
    public int Size;
    /// <summary>Mappings: entries. Sequences: items. Aliases: the anchored record. Scalars: the block indentation.</summary>
    public int Count;
    /// <summary>The row of <see cref="YamlNodeTable.Properties"/> holding the anchor and tag, or -1.</summary>
    public int Props;
    /// <summary>The row of <see cref="YamlNodeTable.Texts"/> holding an alias's name, or -1.</summary>
    public int Text;
    public YamlRecordKind Kind;
    public YamlStyle Style;
    public YamlTokenFlags Flags;
    public YamlChomping Chomping;
}

/// <summary>Anchor, tag and alias facts of a loaded node; only nodes with an anchor or tag have a row.</summary>
internal struct YamlRecordProperties {
    public string? Anchor;
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
    public string[] Texts;
    public int TextCount;
    private readonly string _source;
    /// <summary>Anchored nodes that aliases refer to, by record, once created; aliases find their targets here.</summary>
    private Dictionary<int, YamlNode>? _anchored;
    /// <summary>Shares the text of equal keys across the mappings created from the table.</summary>
    private YamlStringCache? _keys;
    private bool _handedOut;
    private bool _released;

    public YamlNodeTable(string source) {
        _source = source;
        Records = ArrayPool<YamlNodeRecord>.Shared.Rent(Math.Max(16, source.Length / 12));
        Properties = ArrayPool<YamlRecordProperties>.Shared.Rent(8);
        Texts = ArrayPool<string>.Shared.Rent(Math.Max(16, source.Length / 32));
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

    public int AddText(string text) {
        if (TextCount == Texts.Length) Texts = Grow(Texts, TextCount, clear: true);
        Texts[TextCount] = text;
        return TextCount++;
    }

    /// <summary>Returns the pooled arrays when no node object was created from them.</summary>
    public void Release() {
        lock (this) {
            if (_handedOut || _released) return;
            _released = true;
            Array.Clear(Texts, 0, TextCount);
            Array.Clear(Properties, 0, PropertyCount);
            ArrayPool<string>.Shared.Return(Texts);
            ArrayPool<YamlRecordProperties>.Shared.Return(Properties);
            ArrayPool<YamlNodeRecord>.Shared.Return(Records);
            Texts = Array.Empty<string>();
            Properties = Array.Empty<YamlRecordProperties>();
            Records = Array.Empty<YamlNodeRecord>();
            Count = PropertyCount = TextCount = 0;
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

    /// <summary>The scalar's decoded value as a span: the cached key text, the raw source when it is the value, or a decoded copy.</summary>
    public ReadOnlySpan<char> ValueSpan(int record) {
        ref YamlNodeRecord r = ref Records[record];
        if (r.Text >= 0) return Texts[r.Text].AsSpan();
        if (YamlScalarDecoder.IsVerbatim(r.Style, r.Flags)) return _source.AsSpan(r.ValueStart, r.ValueEnd - r.ValueStart);
        return Value(record).AsSpan();
    }

    /// <summary>The scalar's decoded value.</summary>
    public string Value(int record) {
        ref YamlNodeRecord r = ref Records[record];
        if (r.Text >= 0) return Texts[r.Text];
        return YamlScalar.DecodeValue(_source, r.ValueStart, r.ValueEnd, r.Style, r.Flags, r.Chomping, r.Count);
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

    /// <summary>Creates the key and value objects of a loaded mapping on first use.</summary>
    public YamlMappingEntry[] MaterializeEntries(YamlMapping mapping) {
        lock (this) {
            if (mapping._entries is { } existing) return existing;
            ThrowIfReleased();
            YamlDocumentPart part = mapping._part!;
            var entries = new YamlMappingEntry[mapping._count];
            YamlStringCache keys = _keys ??= YamlStringCache.Rent();
            int child = mapping._record + 1;
            for (int i = 0; i < entries.Length; i++) {
                YamlNode key = Create(part, child, mapping, i);
                // Keys are looked up by text, so their text is read now, shared with equal keys of other mappings.
                if (key is YamlScalar scalar && YamlScalarDecoder.IsVerbatim(scalar._style, scalar._flags)) {
                    scalar._value = keys.Get(_source.AsSpan(scalar._valueStart, scalar._valueEnd - scalar._valueStart));
                }
                child += Records[child].Size;
                YamlNode value = Create(part, child, mapping, i);
                child += Records[child].Size;
                entries[i] = new YamlMappingEntry(key, value);
            }
            Volatile.Write(ref mapping._entries, entries);
            return entries;
        }
    }

    /// <summary>Creates the item objects of a loaded sequence on first use.</summary>
    public YamlNode[] MaterializeItems(YamlSequence sequence) {
        lock (this) {
            if (sequence._items is { } existing) return existing;
            ThrowIfReleased();
            YamlDocumentPart part = sequence._part!;
            var items = new YamlNode[sequence._count];
            int child = sequence._record + 1;
            for (int i = 0; i < items.Length; i++) {
                items[i] = Create(part, child, sequence, i);
                child += Records[child].Size;
            }
            Volatile.Write(ref sequence._items, items);
            return items;
        }
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
                    YamlMappingEntry[] entries = mapping.Entries;
                    for (int i = 0; i < mapping._count && next is null; i++) {
                        if (Contains(entries[i].Key, record)) next = entries[i].Key;
                        else if (Contains(entries[i].Value, record)) next = entries[i].Value;
                    }
                } else if (current is YamlSequence sequence) {
                    YamlNode[] items = sequence.Items;
                    for (int i = 0; i < sequence._count && next is null; i++) {
                        if (Contains(items[i], record)) next = items[i];
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
                node = new YamlAlias(part, r.Start, r.End, Texts[r.Text], r.Count);
                break;
        }
        node._record = record;
        node._parent = parent;
        node._index = position;
        if (r.Props >= 0) {
            ref YamlRecordProperties p = ref Properties[r.Props];
            node._properties = new YamlNodeProperties { Anchor = p.Anchor, Tag = p.Tag };
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
