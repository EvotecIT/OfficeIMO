namespace OfficeIMO.Yaml;

/// <summary>
/// A scalar node: plain, quoted or block text. A loaded scalar keeps its source spelling (<c>TRUE</c>, <c>0x1F</c>,
/// <c>~</c>) and decodes its value once, on first use. An omitted node (<c>key:</c> with no value) is an empty plain scalar.
/// </summary>
/// <remarks>
/// Setting a value equal to the current one under the document's schema keeps the source text, so setting
/// <see langword="true"/> on <c>TRUE</c> changes nothing. A different value keeps the scalar's style when that style can
/// represent it, and otherwise takes the most conservative style that does.
/// </remarks>
public sealed class YamlScalar : YamlNode {
    internal int _valueStart;
    internal int _valueEnd;
    internal int _indent;
    internal YamlChomping _chomping;
    internal YamlTokenFlags _flags;
    internal string? _value;

    internal YamlScalar(YamlDocumentPart part, in YamlEvent ev) : base(part, ev.Style, ev.Start, ev.End) {
        _valueStart = ev.ValueStart;
        _valueEnd = ev.ValueEnd;
        _indent = ev.Indent;
        _chomping = ev.Chomping;
        _flags = ev.Flags;
    }

    /// <summary>
    /// Creates a string scalar, or a null scalar (<c>null</c>) when <paramref name="value"/> is <see langword="null"/>.
    /// The style is chosen when the scalar joins a document: plain when no YAML 1.2 core or YAML 1.1 reader could read the
    /// text as anything but a string, otherwise following the quoting its siblings use.
    /// </summary>
    public YamlScalar(string? value) : base(value is null ? YamlStyle.Plain : YamlStyle.None) {
        _value = value ?? "null";
    }

    /// <summary>
    /// Creates a string scalar with a requested style. When it joins a document, a style that cannot represent the value
    /// at its position gives way to the most conservative one that can. A plain style is kept even when the text reads as
    /// another type, so <c>new YamlScalar("on", YamlStyle.Plain)</c> writes <c>on</c>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="style"/> is not a scalar style.</exception>
    public YamlScalar(string value, YamlStyle style) : base(YamlEdit.CheckScalarStyle(style)) {
        _value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Creates an integer scalar.</summary>
    public YamlScalar(long value) : base(YamlStyle.Plain) {
        _value = value.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Creates a floating-point scalar spelled so YAML 1.2 and YAML 1.1 readers read the same float (<c>1.0</c>, <c>.inf</c>).</summary>
    public YamlScalar(double value) : base(YamlStyle.Plain) {
        _value = YamlScalarFormat.Format(value);
    }

    /// <summary>Creates a boolean scalar (<c>true</c> or <c>false</c>).</summary>
    public YamlScalar(bool value) : base(YamlStyle.Plain) {
        _value = YamlScalarFormat.Format(value);
    }

    /// <summary>
    /// Gets the scalar's raw text: for a loaded scalar the content between the quotes with escapes intact, or a block
    /// scalar's content lines with their indentation, without the header; for a scalar set by editing, the text saving
    /// writes for it, with the same parts removed. Use <see cref="YamlNode.AsString"/> for the decoded value.
    /// </summary>
    public string RawText => HasSourceSpan ? Document!._source.Substring(_valueStart, _valueEnd - _valueStart) : YamlModelWriter.RenderRawText(this);

    /// <summary>
    /// Sets a string value, or null with <see langword="null"/>. The style is kept when it can represent the value at
    /// this position; otherwise, or for <see cref="YamlStyle.None"/>, it is chosen as for a new scalar. An equal value
    /// keeps the source text.
    /// </summary>
    /// <returns>This scalar.</returns>
    public YamlScalar SetValue(string? value, YamlStyle style = YamlStyle.None) {
        if (value is null) YamlEdit.AssignScalar(this, "null", YamlScalarKind.Null, YamlStyle.Plain);
        else YamlEdit.AssignScalar(this, value, YamlScalarKind.String, YamlEdit.CheckScalarStyle(style));
        return this;
    }

    /// <summary>Sets an integer value, keeping the source text when it already equals it (<c>0x1F</c> for 31).</summary>
    /// <returns>This scalar.</returns>
    public YamlScalar SetValue(long value) {
        YamlEdit.AssignScalar(this, value.ToString(CultureInfo.InvariantCulture), YamlScalarKind.Integer, YamlStyle.Plain);
        return this;
    }

    /// <summary>Sets a floating-point value, keeping the source text when it already equals it.</summary>
    /// <returns>This scalar.</returns>
    public YamlScalar SetValue(double value) {
        YamlEdit.AssignScalar(this, YamlScalarFormat.Format(value), YamlScalarKind.Float, YamlStyle.Plain);
        return this;
    }

    /// <summary>Sets a boolean value, keeping the source text when it already equals it (<c>TRUE</c>, or <c>yes</c> under YAML 1.1).</summary>
    /// <returns>This scalar.</returns>
    public YamlScalar SetValue(bool value) {
        YamlEdit.AssignScalar(this, YamlScalarFormat.Format(value), YamlScalarKind.Boolean, YamlStyle.Plain);
        return this;
    }

    /// <summary>
    /// Gets what the scalar resolves to under its document's schema: null, boolean, integer, float, string, timestamp, a
    /// YAML 1.1 merge or value key, or <see cref="YamlScalarKind.Custom"/> for a tag no schema resolves. Use it to pick
    /// the <see cref="YamlNode.GetValue{T}"/> type without trying each one.
    /// </summary>
    public YamlScalarKind Kind => TryGetKind(out YamlScalarKind kind) ? kind : YamlScalarKind.Custom;

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
        : YamlScalarDecoder.IsVerbatim(_style, _flags) ? Document!._source.AsSpan(_valueStart, _valueEnd - _valueStart)
        : Value.AsSpan();

    /// <summary>Resolves the scalar's kind as the reader does; returns false for tags no schema resolves.</summary>
    internal bool TryGetKind(out YamlScalarKind kind) {
        string? tag = Tag;
        bool plain = tag is null && _style == YamlStyle.Plain;
        return YamlSchemaResolver.TryGetKind(tag, _style == YamlStyle.None ? YamlStyle.DoubleQuoted : _style, plain ? ValueSpan : default, Schema, out kind);
    }

    /// <summary>The schema used to parse numbers and booleans once the kind is known; failsafe text still parses with tags.</summary>
    internal YamlSchema NumericSchema => Schema == YamlSchema.Failsafe ? YamlSchema.Core : Schema;

    internal bool ResolvesToNull() {
        if (!TryGetKind(out YamlScalarKind kind) || kind != YamlScalarKind.Null) return false;
        return Tag is null || YamlSchemaResolver.IsNull(ValueSpan, NumericSchema);
    }

    /// <summary>Returns whether the scalar is the YAML 1.1 merge key <c>&lt;&lt;</c> in its document.</summary>
    internal bool IsMergeKey() {
        string? tag = Tag;
        if (tag is not null) return tag == YamlSchemaResolver.TagPrefix + "merge";
        if (_style != YamlStyle.Plain || Schema != YamlSchema.Yaml11) return false;
        if (!HasSourceSpan) return _value == "<<";
        string source = Document!._source;
        return _valueEnd - _valueStart == 2 && source[_valueStart] == '<' && source[_valueStart + 1] == '<';
    }

    private string Decode() {
        ReadOnlySpan<char> raw = Document!._source.AsSpan(_valueStart, _valueEnd - _valueStart);
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
