using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>The scalar the pull reader is positioned on.</summary>
internal readonly struct YamlReaderScalar : YamlScalarClassifier.IScalar {
    private readonly YamlReader _reader;

    public YamlReaderScalar(YamlReader reader) {
        _reader = reader;
    }

    public string? Tag => _reader.Tag;
    public YamlStyle Style => _reader.Style;

    /// <summary>
    /// The decoded value. Single-line plain scalars and quoted scalars without escapes are their raw text, which avoids
    /// a string per scalar; anything that needs folding or unescaping is decoded by the engine.
    /// </summary>
    public ReadOnlySpan<char> Text {
        get {
            ReadOnlySpan<char> raw = _reader.ValueSpan;
            YamlStyle style = _reader.Style;
            if (raw.IndexOfAny('\n', '\r') < 0) {
                if (style == YamlStyle.Plain) return raw;
                if (style == YamlStyle.DoubleQuoted && raw.IndexOf('\\') < 0) return raw;
                if (style == YamlStyle.SingleQuoted && raw.IndexOf('\'') < 0) return raw;
            }

            return _reader.GetString().AsSpan();
        }
    }

    public bool IsNull() => _reader.IsNull();
    public bool TryGetBoolean(out bool value) => _reader.TryGetBoolean(out value);
    public bool TryGetInt64(out long value) => _reader.TryGetInt64(out value);
    public bool TryGetDouble(out double value) => _reader.TryGetDouble(out value);
}

/// <summary>A scalar node of a loaded document.</summary>
internal readonly struct YamlNodeScalar : YamlScalarClassifier.IScalar {
    private readonly YamlScalar _node;

    public YamlNodeScalar(YamlScalar node) {
        _node = node;
    }

    public string? Tag => _node.Tag;
    public YamlStyle Style => _node.Style;
    public ReadOnlySpan<char> Text => _node.AsString().AsSpan();
    public bool IsNull() => _node.IsNull;
    public bool TryGetBoolean(out bool value) => _node.TryGetValue(out value);
    public bool TryGetInt64(out long value) => _node.TryGetValue(out value);
    public bool TryGetDouble(out double value) => _node.TryGetValue(out value);
}
