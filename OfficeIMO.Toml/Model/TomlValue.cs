namespace OfficeIMO.Toml;

/// <summary>
/// A string, integer, float, boolean or date-time. A loaded value keeps its source spelling (<c>0xFF</c>, <c>1_000</c>,
/// <c>'literal'</c>) and decodes it on first read.
/// </summary>
public sealed class TomlValue : TomlNode {
    internal TomlNodeKind _kind;
    internal TomlStringStyle _style;
    internal TomlIntegerBase _base;
    /// <summary>The decoded value: string, long, double, bool or <see cref="TomlDateTime"/>. Loaded values fill it on first read.</summary>
    internal object? _value;

    internal TomlValue(TomlDocument document, TomlNodeKind kind) : base(document) {
        _kind = kind;
    }

    /// <inheritdoc />
    public override TomlNodeKind Kind => _kind;

    /// <summary>Gets how a string value is quoted. Other values report <see cref="TomlStringStyle.Bare"/>.</summary>
    public TomlStringStyle StringStyle => _kind == TomlNodeKind.String ? _style : TomlStringStyle.Bare;

    /// <summary>
    /// Gets the decoded value: a <see cref="string"/>, a <see cref="long"/> for integers, a <see cref="double"/> for
    /// floats, a <see cref="bool"/>, or a <see cref="TomlDateTime"/>.
    /// </summary>
    public object Value => _value ??= Decode();

    private object Decode() {
        ReadOnlySpan<char> raw = _document._source.AsSpan(_start, _end - _start);
        switch (_kind) {
            case TomlNodeKind.String:
                return TomlScalars.Decode(raw, _style);
            case TomlNodeKind.Integer:
                TomlScalars.TryParseInteger(raw, out long integer, out _, out _);
                return integer;
            case TomlNodeKind.Float:
                return TomlScalars.ParseFloat(raw);
            case TomlNodeKind.Boolean:
                return raw[0] == 't';
            default:
                TomlScalars.TryParseDateTime(raw, out TomlDateTime dateTime, out _);
                return dateTime;
        }
    }

    /// <summary>
    /// Replaces the value and returns this node. Assigning a value equal to the current one changes nothing, so the source
    /// spelling stays. Otherwise the value is written in the current style when that style can hold it: a hexadecimal
    /// integer stays hexadecimal, a literal string stays literal, a date-time keeps its <c>T</c> or space separator.
    /// </summary>
    /// <param name="value">A string, boolean, integer, floating-point number, <see cref="TomlDateTime"/>,
    /// <see cref="DateTime"/>, <see cref="DateTimeOffset"/>, or on .NET 8 and later <c>DateOnly</c> or <c>TimeOnly</c>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null; TOML has no null.</exception>
    /// <exception cref="ArgumentException"><paramref name="value"/> is a collection or an unsupported type; replace the
    /// node through its parent with <see cref="TomlTable.Set"/> or <see cref="TomlArray.Set"/>.</exception>
    public TomlValue SetValue(object value) {
        if (value is null) throw new ArgumentNullException(nameof(value), "TOML has no null value.");
        if (!TomlValueConverter.TryCreateScalar(value, out TomlNodeKind kind, out object normalized)) {
            throw new ArgumentException("A " + value.GetType().Name + " is not a TOML string, number, boolean or date-time.", nameof(value));
        }
        if (kind == _kind && TomlValueConverter.ScalarEquals(Value, normalized)) return this;
        if (kind != _kind) {
            _kind = kind;
            _base = TomlIntegerBase.Decimal;
        }
        if (kind == TomlNodeKind.String) _style = TomlFormatter.FitStyle((string)normalized, _style);
        if (kind == TomlNodeKind.Integer && (long)normalized < 0) _base = TomlIntegerBase.Decimal;
        _value = normalized;
        if (IsSource) _state |= TomlNodeState.Dirty;
        TomlEditing.MarkChanged(this);
        return this;
    }

    /// <summary>Returns the decoded text of a string, or the TOML spelling of any other value.</summary>
    internal string GetText() {
        if (_kind == TomlNodeKind.String) return (string)Value;
        if (IsSource && (_state & TomlNodeState.Dirty) == 0) return _document._source.Substring(_start, _end - _start);
        var text = new StringBuilder();
        TomlFormatter.AppendScalar(text, this);
        return text.ToString();
    }
}
