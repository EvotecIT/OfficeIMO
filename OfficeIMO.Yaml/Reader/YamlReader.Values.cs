namespace OfficeIMO.Yaml;

public sealed partial class YamlReader {
    private const int TagCacheSize = 8;
    private string?[]? _tagCachePrefix;
    private string?[]? _tagCacheSuffix;
    private string?[]? _tagCacheValue;
    private int _tagCacheNext;
    private YamlStringCache? _strings;

    /// <summary>
    /// Gets the resolved tag of the current node, or <see langword="null"/> when it has none. Shorthand tags are expanded
    /// through the document's <c>%TAG</c> directives (<c>!!str</c> becomes <c>tag:yaml.org,2002:str</c>), percent
    /// escapes are decoded, and the non-specific tag is <c>!</c>. Untagged nodes return <see langword="null"/> even when the
    /// schema would resolve them; use the typed accessors for resolution.
    /// </summary>
    public string? Tag {
        get {
            if (!_event.HasTag || _event.Type is not (YamlTokenType.Scalar or YamlTokenType.SequenceStart or YamlTokenType.MappingStart)) return null;
            string prefix = _event.TagPrefix!;
            ReadOnlySpan<char> suffix = Slice(_event.TagSuffixStart, _event.TagSuffixEnd);
            _tagCachePrefix ??= new string?[TagCacheSize];
            _tagCacheSuffix ??= new string?[TagCacheSize];
            _tagCacheValue ??= new string?[TagCacheSize];
            for (int i = 0; i < TagCacheSize; i++) {
                if ((object?)_tagCachePrefix[i] == prefix && suffix.SequenceEqual(_tagCacheSuffix[i].AsSpan())) return _tagCacheValue[i];
            }
            string value = (_event.TagFlags & YamlTokenFlags.PercentEscapes) != 0 ? prefix + DecodePercent(suffix) : string.Concat(prefix, suffix.ToString());
            int slot = _tagCacheNext++ & (TagCacheSize - 1);
            _tagCachePrefix[slot] = prefix;
            _tagCacheSuffix[slot] = suffix.ToString();
            _tagCacheValue[slot] = value;
            return value;
        }
    }

    /// <summary>
    /// Gets the anchor name of the current node (without <c>&amp;</c>), the name an alias refers to (without <c>*</c>),
    /// or <see langword="null"/>. Repeated names return the same string instance.
    /// </summary>
    public string? Anchor {
        get {
            if (_event.AnchorStart < 0 || _event.Type is not (YamlTokenType.Scalar or YamlTokenType.SequenceStart or YamlTokenType.MappingStart or YamlTokenType.Alias)) return null;
            return Strings.Get(Slice(_event.AnchorStart, _event.AnchorEnd));
        }
    }

    private YamlStringCache Strings => _strings ??= new YamlStringCache();

    /// <summary>
    /// Returns the decoded value of the current scalar (folding, escapes and block chomping applied), the text of the
    /// current comment after <c>#</c>, or the name the current alias refers to. The string is created once per event;
    /// mapping keys that repeat return the same instance.
    /// </summary>
    /// <exception cref="InvalidOperationException">The current event is not a scalar, comment or alias.</exception>
    public string GetString() {
        if (_stringValue is not null) return _stringValue;
        switch (_event.Type) {
            case YamlTokenType.Scalar:
                return _stringValue = DecodeScalar();
            case YamlTokenType.Comment:
                return _stringValue = ValueSpan.ToString();
            case YamlTokenType.Alias:
                return _stringValue = Strings.Get(ValueSpan);
            default:
                throw new InvalidOperationException("GetString can be called only on a scalar, comment or alias event; the current event is " + _event.Type + ".");
        }
    }

    private string DecodeScalar() {
        ReadOnlySpan<char> raw = Slice(_event.ValueStart, _event.ValueEnd);
        if (YamlScalarDecoder.IsVerbatim(_event.Style, _event.Flags)) {
            if (raw.Length == 0) return string.Empty;
            return _event.IsKey ? Strings.Get(raw) : raw.ToString();
        }
        if (_scratch is null || _scratch.Length < raw.Length + 1) {
            if (_scratch is not null) ArrayPool<char>.Shared.Return(_scratch);
            _scratch = ArrayPool<char>.Shared.Rent(Math.Max(raw.Length + 1, 256));
        }
        int length = YamlScalarDecoder.Decode(raw, _event.Style, _event.Chomping, _event.Indent, _scratch);
        return new string(_scratch, 0, length);
    }

    /// <summary>Gets the current scalar's decoded value as a span, without allocating when the source holds it verbatim.</summary>
    internal ReadOnlySpan<char> ScalarValueSpan => ScalarText();

    /// <summary>
    /// Gets whether the current scalar is untagged and cannot be the plain text <c>&lt;&lt;</c>: it is not plain, or its
    /// source is not two characters long. A plain scalar's value is its source on one line, and a value folded from
    /// several lines contains a space, so a merge key's source is exactly <c>&lt;&lt;</c>.
    /// </summary>
    internal bool IsUntaggedNonMergeText => !_event.HasTag && (_event.Style != YamlStyle.Plain || _event.ValueEnd - _event.ValueStart != 2);

    /// <summary>
    /// Returns what <see cref="IsNull"/> returns for the current scalar, checking an untagged scalar's text directly
    /// instead of resolving its kind: only plain <c>null</c> spellings and empty nodes are null without a tag.
    /// </summary>
    internal bool IsNullScalar() {
        if (_event.Type != YamlTokenType.Scalar) return false;
        if (_event.HasTag) return IsNull();
        if (_event.Style != YamlStyle.Plain) return false;
        ReadOnlySpan<char> text = Slice(_event.ValueStart, _event.ValueEnd);
        if (text.Length > 4) return false;
        return YamlSchemaResolver.IsNull(YamlScalarDecoder.IsVerbatim(_event.Style, _event.Flags) ? text : ScalarText(), _documentSchema);
    }

    /// <summary>Gets whether the current node defines an anchor, without creating its name.</summary>
    internal bool HasAnchor => _event.AnchorStart >= 0 && _event.Type is YamlTokenType.Scalar or YamlTokenType.SequenceStart or YamlTokenType.MappingStart;

    /// <summary>Gets the options the reader runs with.</summary>
    internal YamlReaderOptions Options => _options;

    /// <summary>Decodes the current scalar into a span when it needs decoding; otherwise returns the raw value span.</summary>
    private ReadOnlySpan<char> ScalarText() {
        ReadOnlySpan<char> raw = Slice(_event.ValueStart, _event.ValueEnd);
        if (YamlScalarDecoder.IsVerbatim(_event.Style, _event.Flags)) return raw;
        return GetString().AsSpan();
    }

    /// <summary>
    /// Determines how the current scalar resolves: an explicit tag wins, the non-specific tag <c>!</c> and quoted or block
    /// styles mean string, and untagged plain scalars follow the document's schema.
    /// </summary>
    private bool TryGetKind(out YamlScalarKind kind) {
        kind = YamlScalarKind.String;
        if (_event.Type != YamlTokenType.Scalar) return false;
        bool plain = !_event.HasTag && _event.Style == YamlStyle.Plain;
        return YamlSchemaResolver.TryGetKind(_event.HasTag ? Tag : null, _event.Style, plain ? ScalarText() : default, _documentSchema, out kind);
    }

    private YamlSchema NumericSchema => _documentSchema == YamlSchema.Failsafe ? YamlSchema.Core : _documentSchema;

    /// <summary>
    /// Returns whether the current node is null under the active schema: an untagged plain scalar such as <c>null</c>,
    /// <c>~</c> or an empty node (core schema), or a node tagged <c>!!null</c>. Quoted scalars are never null.
    /// </summary>
    public bool IsNull() {
        if (!TryGetKind(out YamlScalarKind kind)) return false;
        if (kind != YamlScalarKind.Null) return false;
        if (!_event.HasTag) return true;
        return YamlSchemaResolver.IsNull(ScalarText(), NumericSchema);
    }

    /// <summary>
    /// Reads the current scalar as a boolean under the active schema (<c>true</c>/<c>false</c> in the 1.2 schemas, plus
    /// <c>yes</c>/<c>no</c>/<c>on</c>/<c>off</c>/<c>y</c>/<c>n</c> under YAML 1.1) or a <c>!!bool</c> tag. Returns
    /// <see langword="false"/> for any other node, including quoted text.
    /// </summary>
    public bool TryGetBoolean(out bool value) {
        value = false;
        if (!TryGetKind(out YamlScalarKind kind) || kind != YamlScalarKind.Boolean) return false;
        return YamlSchemaResolver.TryParseBoolean(ScalarText(), NumericSchema, out value);
    }

    /// <summary>
    /// Reads the current scalar as a 64-bit integer under the active schema (decimal, plus <c>0o</c>/<c>0x</c> in the core
    /// schema and <c>0b</c>, <c>0</c>-octal, underscores and sexagesimal under YAML 1.1) or an <c>!!int</c> tag. Returns
    /// <see langword="false"/> for other nodes and for values outside the <see cref="long"/> range.
    /// </summary>
    public bool TryGetInt64(out long value) {
        value = 0;
        if (!TryGetKind(out YamlScalarKind kind) || kind != YamlScalarKind.Integer) return false;
        return YamlSchemaResolver.TryParseInt64(ScalarText(), NumericSchema, out value);
    }

    /// <summary>
    /// Reads the current scalar as a double: any integer or floating-point value under the active schema, including
    /// <c>.inf</c>, <c>-.inf</c> and <c>.nan</c>, or a node tagged <c>!!int</c> or <c>!!float</c>.
    /// </summary>
    public bool TryGetDouble(out double value) {
        value = 0;
        if (!TryGetKind(out YamlScalarKind kind) || kind is not (YamlScalarKind.Integer or YamlScalarKind.Float)) return false;
        return YamlSchemaResolver.TryParseDouble(ScalarText(), NumericSchema, out value);
    }

    /// <summary>
    /// Reads the current scalar as a decimal: any integer or finite floating-point value under the active schema, or a
    /// node tagged <c>!!int</c> or <c>!!float</c>. Infinity, NaN and out-of-range values return <see langword="false"/>.
    /// </summary>
    public bool TryGetDecimal(out decimal value) {
        value = 0;
        if (!TryGetKind(out YamlScalarKind kind) || kind is not (YamlScalarKind.Integer or YamlScalarKind.Float)) return false;
        return YamlSchemaResolver.TryParseDecimal(ScalarText(), NumericSchema, out value);
    }
}
