namespace OfficeIMO.Yaml;

/// <summary>
/// A scalar node: plain, quoted or block text. It keeps its source spelling (<c>TRUE</c>, <c>0x1F</c>, <c>~</c>) and
/// decodes its value once, on first use. An omitted node (<c>key:</c> with no value) is an empty plain scalar.
/// </summary>
public sealed class YamlScalar : YamlNode {
    internal int _valueStart;
    internal int _valueEnd;
    internal int _indent;
    internal YamlChomping _chomping;
    internal YamlTokenFlags _flags;
    private string? _value;

    internal YamlScalar(YamlDocumentPart part, in YamlEvent ev) : base(part, ev.Style, ev.Start, ev.End) {
        _valueStart = ev.ValueStart;
        _valueEnd = ev.ValueEnd;
        _indent = ev.Indent;
        _chomping = ev.Chomping;
        _flags = ev.Flags;
    }

    /// <summary>
    /// Gets the scalar's raw source text: the content between the quotes with escapes intact, or a block scalar's content
    /// lines with their indentation, without the header. Use <see cref="YamlNode.AsString"/> for the decoded value.
    /// </summary>
    public string RawText => Document._source.Substring(_valueStart, _valueEnd - _valueStart);

    /// <summary>Gets the decoded value, decoding it on first use.</summary>
    internal string Value => _value ??= Decode();

    /// <summary>Stores a value decoded during the load, such as a mapping key from the shared key cache.</summary>
    internal void SetDecodedValue(string value) => _value = value;

    /// <summary>
    /// Gets the decoded value as a span without allocating when the raw text already is the value (single-line plain and
    /// quoted scalars without escapes).
    /// </summary>
    internal ReadOnlySpan<char> ValueSpan =>
        _value is not null ? _value.AsSpan()
        : YamlScalarDecoder.IsVerbatim(_style, _flags) ? Document._source.AsSpan(_valueStart, _valueEnd - _valueStart)
        : Value.AsSpan();

    /// <summary>Resolves the scalar's kind as the reader does; returns false for tags no schema resolves.</summary>
    internal bool TryGetKind(out YamlScalarKind kind) {
        string? tag = Tag;
        bool plain = tag is null && _style == YamlStyle.Plain;
        return YamlSchemaResolver.TryGetKind(tag, _style, plain ? ValueSpan : default, _part.Schema, out kind);
    }

    /// <summary>The schema used to parse numbers and booleans once the kind is known; failsafe text still parses with tags.</summary>
    internal YamlSchema NumericSchema => _part.Schema == YamlSchema.Failsafe ? YamlSchema.Core : _part.Schema;

    internal bool ResolvesToNull() {
        if (!TryGetKind(out YamlScalarKind kind) || kind != YamlScalarKind.Null) return false;
        return Tag is null || YamlSchemaResolver.IsNull(ValueSpan, NumericSchema);
    }

    /// <summary>Returns whether the scalar is the YAML 1.1 merge key <c>&lt;&lt;</c> in its document.</summary>
    internal bool IsMergeKey() {
        string? tag = Tag;
        if (tag is not null) return tag == YamlSchemaResolver.TagPrefix + "merge";
        return _style == YamlStyle.Plain && _part.Schema == YamlSchema.Yaml11
            && _valueEnd - _valueStart == 2 && Document._source[_valueStart] == '<' && Document._source[_valueStart + 1] == '<';
    }

    private string Decode() {
        ReadOnlySpan<char> raw = Document._source.AsSpan(_valueStart, _valueEnd - _valueStart);
        if (YamlScalarDecoder.IsVerbatim(_style, _flags)) return raw.Length == 0 ? string.Empty : raw.ToString();
        char[] buffer = ArrayPool<char>.Shared.Rent(raw.Length + 1);
        try {
            int length = YamlScalarDecoder.Decode(raw, _style, _chomping, _indent, buffer);
            return new string(buffer, 0, length);
        } finally {
            ArrayPool<char>.Shared.Return(buffer);
        }
    }
}
