using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Walks a loaded <see cref="YamlDocument"/> as data events, without recursion. It honors the document's own schema
/// (so <c>%YAML 1.1</c> documents resolve 1.1 timestamps and merge keys), its duplicate-key policy and its comments.
/// </summary>
internal sealed class YamlModelEventReader : DataEventReader {
    private readonly YamlDocument _document;
    private readonly bool _hasDuplicates;
    private readonly int _maxAliasExpansion;
    private readonly Queue<Pending> _pending = new();
    private readonly Stack<Frame> _frames = new();
    private int _part = -1;
    private bool _inPart;
    private long _expanded;
    private YamlNode? _location;
    private YamlSchema _schema;

    private readonly struct Pending {
        public Pending(DataTokenType type, DataScalarKind kind = default, string? text = null) {
            Type = type;
            Kind = kind;
            Text = text;
        }

        public DataTokenType Type { get; }
        public DataScalarKind Kind { get; }
        public string? Text { get; }
    }

    private sealed class Frame {
        public Frame(YamlNode node, bool viaAlias) {
            Node = node;
            ViaAlias = viaAlias;
        }

        public YamlNode Node { get; }
        public bool ViaAlias { get; }
        public List<KeyValuePair<YamlNode, YamlNode>>? Entries;
        public int Index;
    }

    public YamlModelEventReader(YamlDocument document, DataConversionContext context) : base(context) {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _maxAliasExpansion = context.Options.YamlReader?.MaxAliasExpansion ?? 100_000;
        foreach (OfficeDiagnostic diagnostic in document.Diagnostics) {
            if (diagnostic.Code == OfficeDiagnosticCodes.YamlDuplicateKey) _hasDuplicates = true;
            else context.Report(diagnostic.Code, diagnostic.LossKind, diagnostic.Message, diagnostic.Location);
        }

        TrackDuplicateKeys = true;
    }

    public override string? Location => _location is { } node
        ? "line " + node.Start.Line.ToString(CultureInfo.InvariantCulture) + ", column " + node.Start.Column.ToString(CultureInfo.InvariantCulture)
        : null;

    private bool Yaml11 => _schema == YamlSchema.Yaml11;

    protected override bool ReadCore() {
        while (_pending.Count == 0) {
            if (!Advance()) return false;
        }

        Pending next = _pending.Dequeue();
        switch (next.Type) {
            case DataTokenType.PropertyName: SetPropertyName(next.Text.AsSpan()); break;
            case DataTokenType.Scalar: SetScalar(next.Kind, next.Text.AsSpan()); break;
            case DataTokenType.Comment: SetComment(next.Text.AsSpan()); break;
            default: Set(next.Type); break;
        }

        return true;
    }

    private bool Advance() {
        if (!_inPart) {
            if (++_part >= _document.Documents.Count) return false;
            YamlDocumentPart part = _document.Documents[_part];
            _schema = part.Schema;
            _inPart = true;
            _pending.Enqueue(new Pending(DataTokenType.StartDocument));
            if (part.CommentBefore is { } header) {
                foreach (string line in header.Split('\n')) _pending.Enqueue(new Pending(DataTokenType.Comment, text: line.TrimEnd('\r')));
            }

            Begin(part.Root, viaAlias: false);
            return true;
        }

        if (_frames.Count == 0) {
            _inPart = false;
            _pending.Enqueue(new Pending(DataTokenType.EndDocument));
            return true;
        }

        Frame frame = _frames.Peek();
        if (frame.Entries != null) {
            if (frame.Index < frame.Entries.Count) {
                KeyValuePair<YamlNode, YamlNode> entry = frame.Entries[frame.Index++];
                Comments(entry.Key, leading: true);
                _pending.Enqueue(new Pending(DataTokenType.PropertyName, text: KeyText(entry.Key)));
                Begin(entry.Value, viaAlias: false);
                return true;
            }

            End(DataTokenType.EndObject);
            return true;
        }

        var sequence = (YamlSequence)frame.Node;
        if (frame.Index < sequence.Count) {
            YamlNode item = sequence[frame.Index++];
            Comments(item, leading: true);
            Begin(item, viaAlias: false);
            return true;
        }

        End(DataTokenType.EndArray);
        return true;
    }

    private void End(DataTokenType type) {
        Frame frame = _frames.Pop();
        _pending.Enqueue(new Pending(type));
        Comments(frame.Node, leading: false);
    }

    private void Begin(YamlNode node, bool viaAlias) {
        if (node is YamlAlias alias) {
            Context.Report(OfficeDiagnosticCodes.YamlAliasExpanded, OfficeConversionLossKind.None,
                "YAML aliases were expanded into copies of their anchored values; the data is unchanged.");
            Begin(alias.Target, viaAlias: true);
            return;
        }

        _location = node;
        if (viaAlias || (_frames.Count > 0 && _frames.Peek().ViaAlias)) {
            if (++_expanded > _maxAliasExpansion) {
                throw new InvalidDataException("Expanding aliases visited more than " + _maxAliasExpansion.ToString(CultureInfo.InvariantCulture)
                    + " nodes (MaxAliasExpansion)" + At() + ".");
            }

            viaAlias = true;
        }

        switch (node) {
            case YamlMapping mapping:
                YamlScalarClassifier.CheckCollectionTag(mapping.Tag, Context);
                _pending.Enqueue(new Pending(DataTokenType.StartObject));
                _frames.Push(new Frame(mapping, viaAlias) { Entries = EffectiveEntries(mapping, 0) });
                break;
            case YamlSequence sequence:
                YamlScalarClassifier.CheckCollectionTag(sequence.Tag, Context);
                _pending.Enqueue(new Pending(DataTokenType.StartArray));
                _frames.Push(new Frame(sequence, viaAlias));
                break;
            case YamlScalar scalar:
                var view = new YamlNodeScalar(scalar);
                DataScalarKind kind = YamlScalarClassifier.Classify(ref view, Yaml11, Context, out ReadOnlySpan<char> text);
                _pending.Enqueue(new Pending(DataTokenType.Scalar, kind, text.ToString()));
                Comments(scalar, leading: false);
                break;
        }
    }

    private void Comments(YamlNode node, bool leading) {
        string? text = leading ? node.CommentBefore : node.Comment;
        if (text == null) return;
        foreach (string line in text.Split('\n')) _pending.Enqueue(new Pending(DataTokenType.Comment, text: line.TrimEnd('\r')));
    }

    private string KeyText(YamlNode key) {
        YamlNode resolved = key is YamlAlias alias ? alias.Target : key;
        if (resolved is YamlScalar scalar) {
            var view = new YamlNodeScalar(scalar);
            return YamlScalarClassifier.KeyText(ref view, Yaml11, Context).ToString();
        }

        Context.Report(OfficeDiagnosticCodes.YamlComplexKey, OfficeConversionLossKind.Approximation,
            "Mappings or sequences used as mapping keys were written as their flow YAML text.");
        return FlowText(resolved);
    }

    private static string FlowText(YamlNode node) {
        string text = node.ToString();
        var builder = new StringBuilder(text.Length);
        bool space = false;
        foreach (char c in text) {
            if (char.IsWhiteSpace(c)) {
                space = builder.Length > 0;
                continue;
            }

            if (space) builder.Append(' ');
            space = false;
            builder.Append(c);
        }

        return builder.ToString();
    }

    // The document decides: YAML 1.1, or the core schema unless the document was loaded strictly.
    private static bool IsMergeKey(YamlNode key) => key is YamlScalar scalar && scalar.Kind == YamlScalarKind.Merge;

    /// <summary>
    /// The entries a reader sees: merge keys replaced by the merged mappings' entries that are not set explicitly, and,
    /// when the document was loaded with a first- or last-wins policy, only the winning entry of each duplicated key.
    /// </summary>
    private List<KeyValuePair<YamlNode, YamlNode>> EffectiveEntries(YamlMapping mapping, int depth) {
        if (depth > 64) throw new InvalidDataException("Merge keys nest more than 64 levels" + At() + ".");
        var entries = new List<KeyValuePair<YamlNode, YamlNode>>(mapping.Count);
        List<KeyValuePair<YamlNode, YamlNode>>? merged = null;
        int mergeIndex = 0;
        for (int i = 0; i < mapping.Count; i++) {
            YamlMappingEntry entry = mapping[i];
            if (IsMergeKey(entry.Key)) {
                if (merged == null) {
                    merged = new List<KeyValuePair<YamlNode, YamlNode>>();
                    mergeIndex = entries.Count;
                }

                YamlNode value = entry.Value is YamlAlias alias ? alias.Target : entry.Value;
                if (value is YamlSequence sources) {
                    foreach (YamlNode source in sources) merged.AddRange(EffectiveEntries(source.AsMapping(), depth + 1));
                } else {
                    merged.AddRange(EffectiveEntries(value.AsMapping(), depth + 1));
                }

                continue;
            }

            if (_hasDuplicates && entry.Key is YamlScalar keyScalar && mapping.TryGetValue(keyScalar.AsString(), out YamlNode? winner)
                && !ReferenceEquals(winner, entry.Value)) {
                Context.Report(OfficeDiagnosticCodes.DuplicateKey, OfficeConversionLossKind.Omission,
                    "Duplicate mapping keys were resolved by the document's duplicate-key policy; the losing values were dropped.",
                    "line " + entry.Key.Start.Line.ToString(CultureInfo.InvariantCulture));
                continue;
            }

            entries.Add(new KeyValuePair<YamlNode, YamlNode>(entry.Key, entry.Value));
        }

        if (merged != null) {
            var explicitKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (KeyValuePair<YamlNode, YamlNode> entry in entries) {
                if (entry.Key is YamlScalar scalar) explicitKeys.Add(scalar.AsString());
            }

            var accepted = new List<KeyValuePair<YamlNode, YamlNode>>(merged.Count);
            foreach (KeyValuePair<YamlNode, YamlNode> entry in merged) {
                if (entry.Key is not YamlScalar scalar || explicitKeys.Add(scalar.AsString())) accepted.Add(entry);
            }

            entries.InsertRange(mergeIndex, accepted);
        }

        return entries;
    }
}
