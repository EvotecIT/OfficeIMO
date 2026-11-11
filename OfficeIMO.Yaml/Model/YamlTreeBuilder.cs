namespace OfficeIMO.Yaml;

/// <summary>
/// Builds the document's <see cref="YamlNodeTable"/> from <see cref="YamlReader"/> events in one pass without recursion:
/// one record per node, in document order, and no node objects. The builder also enforces the per-document node and
/// alias-expansion limits, resolves aliases to their anchored records and checks mappings for duplicate keys.
/// </summary>
internal sealed class YamlTreeBuilder {
    /// <summary>Saturation point for expansion sizes, far above any limit and safe from overflow when added.</summary>
    private const long ExpansionCap = long.MaxValue / 4;
    private const int LinearDuplicateLimit = 8;
    private const char Marker = '\u0001';
    private const string NullIdentity = "\u0001null";
    private const string TrueIdentity = "\u0001true";
    private const string FalseIdentity = "\u0001false";

    private readonly YamlDocument _document;
    private readonly YamlLoadOptions _options;
    private readonly YamlNodeTable _table;
    private readonly string _source;
    private readonly int _maxNodes;
    private Frame[] _frames = ArrayPool<Frame>.Shared.Rent(16);
    private int _frameCount;
    /// <summary>The anchors defined so far in the current document, by name, in open addressing over their source spans.</summary>
    private AnchorSlot[] _anchors = Array.Empty<AnchorSlot>();
    private int _anchorCount;
    private YamlDocumentPart? _part;
    private int _partRoot;
    private int _documentNodes;
    private long _expansion;
    /// <summary>The identities of the keys of the open mappings, innermost last; each frame's keys start at its <see cref="Frame.KeyBase"/>.</summary>
    private KeyIdentity[] _keys = ArrayPool<KeyIdentity>.Shared.Rent(16);
    private int _keyCount;
    private int[] _buckets = ArrayPool<int>.Shared.Rent(16);
    private int _findingsTaken;
    /// <summary>Whether the current part's merge keys still need their one note: they merge under the core schema only because the load is not strict.</summary>
    private bool _noteMergeKeys;

    private struct AnchorSlot {
        public bool Used;
        public int Hash;
        public int Start;
        public int Length;
        public int Record;
    }

    private struct Frame {
        public int Record;
        public int Children;
        public long Expanded;
        /// <summary>Where the mapping's key identities start in <see cref="_keys"/>.</summary>
        public int KeyBase;
        public bool Mapping;
        /// <summary>A '&lt;&lt;' key was added, so the mapping's merge keys may need the part's note.</summary>
        public bool MergeKey;
    }

    public YamlTreeBuilder(YamlDocument document, YamlLoadOptions options, YamlNodeTable table) {
        _document = document;
        _options = options;
        _table = table;
        _source = document._source;
        _maxNodes = options.MaxNodes;
    }

    /// <summary>Reads every event into the table, then returns the builder's pooled scratch arrays.</summary>
    public void Run(YamlReader reader) {
        try {
            Build(reader);
        } finally {
            ArrayPool<Frame>.Shared.Return(_frames);
            // Closed mappings clear their identities; after a failure the open ones are cleared here.
            Array.Clear(_keys, 0, _keyCount);
            ArrayPool<KeyIdentity>.Shared.Return(_keys);
            ArrayPool<int>.Shared.Return(_buckets);
        }
    }

    private void Build(YamlReader reader) {
        while (reader.Read()) {
            if (reader.LenientFindings is { } findings && findings.Count > _findingsTaken) TakeFindings(findings);
            ref readonly YamlEvent ev = ref reader.CurrentEvent;
            switch (ev.Type) {
                case YamlTokenType.DocumentStart:
                    _part = new YamlDocumentPart(_document, reader.DocumentSchema, reader.DocumentVersion, reader.CopyTagDirectives(), !ev.Implicit, ev.Start, ev.End);
                    if (_anchorCount > 0) {
                        Array.Clear(_anchors, 0, _anchors.Length);
                        _anchorCount = 0;
                    }
                    _noteMergeKeys = _part.Schema == YamlSchema.Core && _part.MergeKeys;
                    _expansion = 0;
                    _documentNodes = 0;
                    _partRoot = _table.Count;
                    break;
                case YamlTokenType.DocumentEnd:
                    _part!._table = _table;
                    _part._rootRecord = _partRoot;
                    _part._end = ev.End;
                    _part.HasEndMarker = !ev.Implicit;
                    _document.AddPart(_part);
                    break;
                case YamlTokenType.Scalar: {
                    int record = _table.Add();
                    ref YamlNodeRecord r = ref _table.Records[record];
                    r.Kind = YamlRecordKind.Scalar;
                    r.Style = ev.Style;
                    r.Flags = ev.Flags;
                    r.Chomping = ev.Chomping;
                    r.Start = ev.Start;
                    r.End = ev.End;
                    r.ValueStart = ev.ValueStart;
                    r.ValueEnd = ev.ValueEnd;
                    r.Count = ev.Indent;
                    r.Size = 1;
                    r.Props = -1;
                    bool properties = ev.AnchorStart >= 0 || ev.HasTag;
                    if (properties) SetProperties(record, reader, in ev, open: false);
                    if (++_documentNodes > _maxNodes) ThrowNodes(in ev);
                    // A merge key is '<<' or tagged; other keys skip the check.
                    if (ev.IsKey && _noteMergeKeys && _frameCount > 0 && (ev.HasTag || (ev.ValueEnd - ev.ValueStart == 2 && _source[ev.ValueStart] == '<'))
                        && IsMergeKey(record)) {
                        _frames[_frameCount - 1].MergeKey = true;
                    }
                    // A key without properties whose text is its value and reads as a string under every schema is
                    // identified by its source text here, without reading its record again.
                    if (!properties && _frameCount > 0 && _frames[_frameCount - 1].Mapping && (_frames[_frameCount - 1].Children & 1) == 0
                        && YamlScalarDecoder.IsVerbatim(ev.Style, ev.Flags)
                        && (ev.Style != YamlStyle.Plain || IsAlwaysString(_source.AsSpan(ev.ValueStart, ev.ValueEnd - ev.ValueStart)))) {
                        AddKeyToParent(new KeyIdentity { Class = StringKey, Start = ev.ValueStart, Length = ev.ValueEnd - ev.ValueStart, Record = record });
                    } else {
                        AddToParent(record, 1);
                    }
                    break;
                }
                case YamlTokenType.Alias:
                    AddAlias(in ev);
                    break;
                case YamlTokenType.MappingStart:
                    Open(YamlRecordKind.Mapping, reader, in ev);
                    break;
                case YamlTokenType.SequenceStart:
                    Open(YamlRecordKind.Sequence, reader, in ev);
                    break;
                case YamlTokenType.MappingEnd:
                case YamlTokenType.SequenceEnd:
                    Close(in ev);
                    break;
            }
        }
    }

    private void Open(YamlRecordKind kind, YamlReader reader, in YamlEvent ev) {
        int record = _table.Add();
        ref YamlNodeRecord r = ref _table.Records[record];
        r.Kind = kind;
        r.Style = ev.Style;
        r.Flags = 0;
        r.Chomping = 0;
        r.Start = ev.Start;
        r.End = ev.End;
        r.ValueStart = ev.End;
        r.ValueEnd = ev.End;
        r.Count = 0;
        r.Size = 1;
        r.Props = -1;
        SetProperties(record, reader, in ev, open: true);
        Count(in ev);
        if (_frameCount == _frames.Length) _frames = Grow(_frames, _frameCount, _frameCount * 2);
        _frames[_frameCount++] = new Frame { Record = record, Expanded = 1, KeyBase = _keyCount, Mapping = kind == YamlRecordKind.Mapping };
    }

    private void Close(in YamlEvent ev) {
        Frame frame = _frames[--_frameCount];
        ref YamlNodeRecord r = ref _table.Records[frame.Record];
        r.Size = _table.Count - frame.Record;
        r.End = ev.End;
        if (frame.Mapping) {
            r.Count = frame.Children / 2;
            try {
                if (_noteMergeKeys && frame.MergeKey) NoteMergeKeys(frame.KeyBase, r.Count);
                if (r.Count > 1) CheckDuplicates(frame.KeyBase, r.Count);
            } finally {
                // The identities hold strings; the array keeps none past their mapping.
                Array.Clear(_keys, frame.KeyBase, _keyCount - frame.KeyBase);
                _keyCount = frame.KeyBase;
            }
        } else {
            r.Count = frame.Children;
        }
        if (r.Props >= 0) {
            ref YamlRecordProperties properties = ref _table.Properties[r.Props];
            properties.ExpandedSize = frame.Expanded;
            properties.Open = false;
        }
        AddToParent(frame.Record, frame.Expanded);
    }

    private void AddAlias(in YamlEvent ev) {
        int nameStart = ev.AnchorStart, nameLength = ev.AnchorEnd - ev.AnchorStart;
        int target = FindAnchor(nameStart, nameLength);
        if (target < 0) {
            throw new YamlException(YamlErrorCodes.UndefinedAlias, "The alias *" + _source.Substring(nameStart, nameLength) + " has no anchor before it in its document.", Position(in ev));
        }
        int limit = _options.MaxAliasExpansion;
        ref YamlRecordProperties properties = ref _table.Properties[_table.Records[target].Props];
        // An alias inside the collection it refers to makes the resolved graph infinite.
        if (properties.Open) {
            throw new YamlException(YamlErrorCodes.RecursiveAlias, "The alias *" + _source.Substring(nameStart, nameLength)
                + " is inside the collection its anchor names, which would make the document infinite.", Position(in ev));
        }
        long size = properties.ExpandedSize;
        _expansion = Saturate(_expansion + size);
        if (_expansion > limit) throw new YamlLimitExceededException(nameof(YamlLoadOptions.MaxAliasExpansion), limit, Position(in ev));
        Count(in ev);
        properties.Aliased = true;
        int record = _table.Add();
        ref YamlNodeRecord r = ref _table.Records[record];
        r.Kind = YamlRecordKind.Alias;
        r.Style = YamlStyle.None;
        r.Flags = 0;
        r.Chomping = 0;
        r.Start = ev.Start;
        r.End = ev.End;
        r.ValueStart = nameStart;
        r.ValueEnd = ev.AnchorEnd;
        r.Count = target;
        r.Size = 1;
        r.Props = -1;
        AddToParent(record, size);
    }

    private void SetProperties(int record, YamlReader reader, in YamlEvent ev, bool open) {
        if (ev.AnchorStart < 0 && !ev.HasTag) return;
        int row = _table.AddProperties();
        _table.Records[record].Props = row;
        ref YamlRecordProperties properties = ref _table.Properties[row];
        properties.Tag = reader.Tag;
        properties.ExpandedSize = 1;
        properties.AnchorStart = -1;
        if (ev.AnchorStart < 0) return;
        properties.AnchorStart = ev.AnchorStart;
        properties.AnchorLength = ev.AnchorEnd - ev.AnchorStart;
        properties.Open = open;
        DefineAnchor(ev.AnchorStart, ev.AnchorEnd - ev.AnchorStart, record);
    }

    /// <summary>Makes an anchor name refer to a record; a later anchor with the same name replaces the earlier one.</summary>
    private void DefineAnchor(int start, int length, int record) {
        if ((_anchorCount + 1) * 2 > _anchors.Length) GrowAnchors();
        int hash = YamlNodeTable.HashText(_source.AsSpan(start, length));
        int mask = _anchors.Length - 1;
        for (int slot = hash & mask; ; slot = (slot + 1) & mask) {
            ref AnchorSlot anchor = ref _anchors[slot];
            if (!anchor.Used) {
                anchor = new AnchorSlot { Used = true, Hash = hash, Start = start, Length = length, Record = record };
                _anchorCount++;
                return;
            }
            if (anchor.Hash == hash && _source.AsSpan(anchor.Start, anchor.Length).SequenceEqual(_source.AsSpan(start, length))) {
                anchor.Record = record;
                return;
            }
        }
    }

    /// <summary>Returns the record the anchor name at a source span refers to, or -1.</summary>
    private int FindAnchor(int start, int length) {
        if (_anchorCount == 0) return -1;
        int hash = YamlNodeTable.HashText(_source.AsSpan(start, length));
        int mask = _anchors.Length - 1;
        for (int slot = hash & mask; _anchors[slot].Used; slot = (slot + 1) & mask) {
            ref AnchorSlot anchor = ref _anchors[slot];
            if (anchor.Hash == hash && _source.AsSpan(anchor.Start, anchor.Length).SequenceEqual(_source.AsSpan(start, length))) return anchor.Record;
        }
        return -1;
    }

    private void GrowAnchors() {
        AnchorSlot[] old = _anchors;
        _anchors = new AnchorSlot[Math.Max(16, old.Length * 2)];
        int mask = _anchors.Length - 1;
        foreach (AnchorSlot anchor in old) {
            if (!anchor.Used) continue;
            int slot = anchor.Hash & mask;
            while (_anchors[slot].Used) slot = (slot + 1) & mask;
            _anchors[slot] = anchor;
        }
    }

    private void Count(in YamlEvent ev) {
        if (++_documentNodes > _maxNodes) ThrowNodes(in ev);
    }

    private void ThrowNodes(in YamlEvent ev) =>
        throw new YamlLimitExceededException(nameof(YamlLoadOptions.MaxNodes), _maxNodes, Position(in ev));

    /// <summary>Counts a scalar key already identified as a child of the innermost open mapping, as <see cref="AddToParent"/> does.</summary>
    private void AddKeyToParent(in KeyIdentity identity) {
        ref Frame parent = ref _frames[_frameCount - 1];
        if (_keyCount == _keys.Length) _keys = Grow(_keys, _keyCount, _keyCount * 2, clear: true);
        ref KeyIdentity key = ref _keys[_keyCount++];
        key = identity;
        key.Signature = Signature(_source.AsSpan(identity.Start, identity.Length));
        parent.Children++;
        parent.Expanded = Saturate(parent.Expanded + 1);
    }

    /// <summary>
    /// Counts a finished node as a child of the innermost open collection. A mapping's key gets its identity now, while
    /// its record and text are fresh, for the duplicate check when the mapping closes.
    /// </summary>
    private void AddToParent(int record, long expanded) {
        if (_frameCount == 0) return;
        ref Frame parent = ref _frames[_frameCount - 1];
        if (parent.Mapping && (parent.Children & 1) == 0) {
            if (_keyCount == _keys.Length) _keys = Grow(_keys, _keyCount, _keyCount * 2, clear: true);
            ref KeyIdentity identity = ref _keys[_keyCount++];
            identity = Identify(record);
            identity.Record = record;
            identity.Signature = Signature(Span(in identity));
        }
        parent.Children++;
        parent.Expanded = Saturate(parent.Expanded + expanded);
    }

    /// <summary>Replaces a pooled scratch array with a larger one, keeping its first <paramref name="keep"/> elements.</summary>
    private static T[] Grow<T>(T[] array, int keep, int size, bool clear = false) {
        T[] larger = ArrayPool<T>.Shared.Rent(size);
        Array.Copy(array, larger, keep);
        ArrayPool<T>.Shared.Return(array, clear);
        return larger;
    }

    private static long Saturate(long value) => value > ExpansionCap ? ExpansionCap : value;

    private static YamlPosition Position(in YamlEvent ev) => new YamlPosition(ev.Start, ev.StartLine + 1, ev.StartColumn + 1);

    /// <summary>
    /// Finds keys that are equal as YAML values (specification section 3.2.1.3). Each key gets an identity: its decoded
    /// text for strings, read from the source without a copy when the text is the value, or a canonical form for other
    /// kinds. Equal identities are confirmed by <see cref="Equal"/>. Small mappings compare identities pairwise, larger
    /// ones through a reused hash table, so checking keys allocates nothing in the common case.
    /// </summary>
    private void CheckDuplicates(int first, int count) {
        KeyIdentity[] keys = _keys;
        if (count <= LinearDuplicateLimit) {
            for (int i = 1; i < count; i++) {
                for (int j = 0; j < i; j++) {
                    if (Same(in keys[first + i], in keys[first + j]) && Equal(keys[first + j].Record, keys[first + i].Record)) {
                        ReportDuplicate(keys[first + i].Record, keys[first + j].Record);
                        break;
                    }
                }
            }
            return;
        }
        int size = 16;
        while (size < count * 2) size <<= 1;
        if (_buckets.Length < size) _buckets = Grow(_buckets, 0, size);
        int[] buckets = _buckets;
        Array.Clear(buckets, 0, size);
        int mask = size - 1;
        for (int i = 0; i < count; i++) keys[first + i].Hash = Hash(in keys[first + i]);
        for (int i = 0; i < count; i++) {
            ref KeyIdentity key = ref keys[first + i];
            int slot = key.Hash & mask;
            bool seen = false;
            for (int known; (known = buckets[slot]) != 0; slot = (slot + 1) & mask) {
                if (Same(in keys[first + known - 1], in key)) {
                    seen = true;
                    break;
                }
            }
            if (seen || key.Class == CollectionKey) {
                // Collection keys and identity collisions are rare: confirm against every earlier key.
                for (int j = 0; j < i; j++) {
                    if (Same(in keys[first + j], in key) && Equal(keys[first + j].Record, key.Record)) {
                        ReportDuplicate(key.Record, keys[first + j].Record);
                        break;
                    }
                }
            }
            if (!seen) buckets[slot] = i + 1;
        }
    }

    private const byte CollectionKey = 0;
    private const byte StringKey = 1;
    private const byte OtherKey = 2;

    /// <summary>
    /// A key's identity for duplicate checks: equal for keys equal as values, though unequal keys may share one. String
    /// keys refer to their text (a source range or a decoded copy); other kinds hold a canonical form.
    /// </summary>
    private struct KeyIdentity {
        public byte Class;
        public string? Text;
        public int Start;
        public int Length;
        /// <summary>A hash of the class and text, compared before the text.</summary>
        public int Hash;
        /// <summary>The text's first and last characters, compared before the text.</summary>
        public int Signature;
        public int Record;
    }

    private ReadOnlySpan<char> Span(in KeyIdentity identity) => identity.Text is { } text ? text.AsSpan() : _source.AsSpan(identity.Start, identity.Length);

    /// <summary>Compares hashed identities, the hashes first.</summary>
    private bool Same(in KeyIdentity a, in KeyIdentity b) =>
        a.Length == b.Length && a.Signature == b.Signature && a.Class == b.Class && (a.Class == CollectionKey || Span(in a).SequenceEqual(Span(in b)));

    private static int Signature(ReadOnlySpan<char> text) => text.Length == 0 ? 0 : (text[0] << 16) | text[text.Length - 1];

    private int Hash(in KeyIdentity identity) => YamlNodeTable.HashText(Span(in identity)) ^ (identity.Class * 0x45D9F3B);

    /// <summary>Resolves a scalar record's kind as the reader does; returns false for tags no schema resolves.</summary>
    private bool TryGetKind(int record, out YamlScalarKind kind) {
        string? tag = _table.Tag(record);
        YamlStyle style = _table.Records[record].Style;
        bool plain = tag is null && style == YamlStyle.Plain;
        YamlSchema schema = _part!.Schema;
        bool resolved = YamlSchemaResolver.TryGetKind(tag, style, plain ? _table.ValueSpan(record) : default, schema, out kind);
        // Outside strict mode the core schema reads '<<' as a merge key too, as YAML 1.1 does.
        if (plain && kind == YamlScalarKind.String && schema == YamlSchema.Core && _part.MergeKeys && IsMergeKey(record)) kind = YamlScalarKind.Merge;
        return resolved;
    }

    private bool IsMergeKey(int record) => _table.IsMergeKey(record, _part!.MergeKeys);

    /// <summary>
    /// Returns a key's identity; equal identities correspond to equal values of <see cref="YamlKeyEquality.Identity"/> for
    /// the nodes the records become.
    /// </summary>
    private KeyIdentity Identify(int key) {
        int record = _table.Resolve(key);
        ref YamlNodeRecord r = ref _table.Records[record];
        if (r.Kind != YamlRecordKind.Scalar) return default;
        // Untagged quoted keys, and plain keys no schema reads as anything but a string, are their text.
        if (r.Props < 0 && (r.Style != YamlStyle.Plain || IsAlwaysString(_source.AsSpan(r.ValueStart, r.ValueEnd - r.ValueStart)))) return StringIdentity(record);
        if (!TryGetKind(record, out YamlScalarKind kind)) return Other(_table.Tag(record) + Marker + _table.Value(record));
        YamlSchema schema = _part!.Schema == YamlSchema.Failsafe ? YamlSchema.Core : _part.Schema;
        switch (kind) {
            case YamlScalarKind.Null:
                return Other(NullIdentity);
            case YamlScalarKind.Boolean:
                if (YamlSchemaResolver.TryParseBoolean(_table.ValueSpan(record), schema, out bool flag)) return Other(flag ? TrueIdentity : FalseIdentity);
                break;
            case YamlScalarKind.Integer:
                if (YamlSchemaResolver.TryParseInt64(_table.ValueSpan(record), schema, out long integer)) return Other(Marker + integer.ToString(CultureInfo.InvariantCulture));
                break;
            case YamlScalarKind.Float:
                if (YamlSchemaResolver.TryParseDouble(_table.ValueSpan(record), schema, out double number)) return Other(Marker + number.ToString("R", CultureInfo.InvariantCulture));
                break;
            case YamlScalarKind.String:
                return StringIdentity(record);
        }
        return Other(Marker + _table.Value(record));

        static KeyIdentity Other(string text) => new KeyIdentity { Class = OtherKey, Text = text, Length = text.Length };
    }

    private KeyIdentity StringIdentity(int record) {
        ref YamlNodeRecord r = ref _table.Records[record];
        if (YamlScalarDecoder.IsVerbatim(r.Style, r.Flags)) {
            return new KeyIdentity { Class = StringKey, Start = r.ValueStart, Length = r.ValueEnd - r.ValueStart };
        }
        string value = _table.Value(record);
        return new KeyIdentity { Class = StringKey, Text = value, Length = value.Length };
    }

    /// <summary>
    /// Returns whether a plain scalar reads as a string under every schema. Numbers, timestamps, <c>~</c> and the YAML
    /// 1.1 <c>&lt;&lt;</c> and <c>=</c> keys start with a digit, sign, dot or symbol, and the nulls and booleans that
    /// start with a letter are the words checked here in any letter case, a superset of the spellings schemas accept.
    /// </summary>
    private static bool IsAlwaysString(ReadOnlySpan<char> text) {
        if (text.Length == 0) return false;
        char first = text[0];
        if (!(first is >= 'a' and <= 'z' or >= 'A' and <= 'Z')) return first is not ((>= '0' and <= '9') or '+' or '-' or '.' or '~' or '<' or '=');
        // The words start with these letters and are at most five characters long.
        if (text.Length > 5 || (first | 0x20) is not ('n' or 'o' or 'y' or 't' or 'f')) return true;
        return text.Length switch {
            1 => (first | 0x20) is not ('y' or 'n'),
            2 => !Is(text, "no") && !Is(text, "on"),
            3 => !Is(text, "yes") && !Is(text, "off"),
            4 => !Is(text, "null") && !Is(text, "true"),
            5 => !Is(text, "false"),
            _ => true
        };

        // ASCII case folding: only a letter's two cases fold to a lowercase letter, and the words are lowercase.
        static bool Is(ReadOnlySpan<char> text, string word) {
            for (int i = 0; i < word.Length; i++) {
                if ((text[i] | 0x20) != word[i]) return false;
            }
            return true;
        }
    }

    /// <summary>Compares two key records as values, as <see cref="YamlKeyEquality.Equal"/> compares nodes.</summary>
    private bool Equal(int left, int right) {
        int a = _table.Resolve(left), b = _table.Resolve(right);
        if (a == b) return true;
        ref YamlNodeRecord ra = ref _table.Records[a];
        ref YamlNodeRecord rb = ref _table.Records[b];
        if (ra.Kind != rb.Kind) return false;
        switch (ra.Kind) {
            case YamlRecordKind.Scalar:
                bool ka = TryGetKind(a, out YamlScalarKind kindA), kb = TryGetKind(b, out YamlScalarKind kindB);
                if (ka != kb) return false;
                if (!ka) return _table.Tag(a) == _table.Tag(b) && _table.ValueSpan(a).SequenceEqual(_table.ValueSpan(b));
                // Repeated YAML 1.1 merge keys are merge instructions, not duplicate entries.
                if (kindA != kindB || kindA == YamlScalarKind.Merge) return false;
                // Integers and floats share the "\u0001" prefix; an equal identity with different kinds is not a duplicate.
                KeyIdentity ia = Identify(a), ib = Identify(b);
                return ia.Class == ib.Class && Span(in ia).SequenceEqual(Span(in ib));
            case YamlRecordKind.Mapping:
            case YamlRecordKind.Sequence:
                int children = ra.Kind == YamlRecordKind.Mapping ? ra.Count * 2 : ra.Count;
                if (ra.Count != rb.Count || _table.Tag(a) != _table.Tag(b)) return false;
                for (int i = 0, x = a + 1, y = b + 1; i < children; i++, x = _table.Next(x), y = _table.Next(y)) {
                    if (!Equal(x, y)) return false;
                }
                return true;
            default:
                return false;
        }
    }

    private void TakeFindings(List<YamlLenientFinding> findings) {
        for (; _findingsTaken < findings.Count; _findingsTaken++) _document.AddDiagnostic(findings[_findingsTaken].ToDiagnostic(_source));
    }

    /// <summary>Notes, once per document part, that '&lt;&lt;' keys merge although no %YAML 1.1 directive asks for it.</summary>
    private void NoteMergeKeys(int first, int count) {
        for (int i = 0; i < count; i++) {
            int key = _keys[first + i].Record;
            if (_table.Records[key].Kind != YamlRecordKind.Scalar || !IsMergeKey(key)) continue;
            _noteMergeKeys = false;
            YamlPosition position = _document.GetPosition(_table.Records[key].Start);
            _document.AddDiagnostic(new OfficeDiagnostic(OfficeDiagnosticCodes.YamlMergeKeyWithoutDirective, OfficeDiagnosticSeverity.Information,
                "The document uses '<<' merge keys, such as the one at " + position + ", without a %YAML 1.1 directive; they were resolved as YAML 1.1 and common tools resolve them. Load with Strict to read '<<' as an ordinary key.",
                position.ToString(), "YAML"));
            return;
        }
    }

    private void ReportDuplicate(int key, int first) {
        int resolved = _table.Resolve(key);
        ref YamlNodeRecord r = ref _table.Records[key];
        string text = _table.Records[resolved].Kind == YamlRecordKind.Scalar ? _table.Value(resolved) : _source.Substring(r.Start, r.End - r.Start);
        if (text.Length > 64) text = text.Substring(0, 61) + "...";
        YamlPosition position = _document.GetPosition(r.Start);
        string message = "The mapping key '" + text + "' at " + position + " repeats the key at " + _document.GetPosition(_table.Records[first].Start) + ".";
        if (_options.EffectiveDuplicateKeys == YamlDuplicateKeyPolicy.Error) throw new YamlException(YamlErrorCodes.DuplicateKey, message, position);
        _document.AddDiagnostic(new OfficeDiagnostic(OfficeDiagnosticCodes.YamlDuplicateKey, OfficeDiagnosticSeverity.Warning,
            message + YamlExcerpt.Format(_source, r.Start), position.ToString(), "YAML"));
    }
}
