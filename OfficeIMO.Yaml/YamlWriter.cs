using OfficeIMO.SharedSource.IO;

namespace OfficeIMO.Yaml;

/// <summary>
/// A forward-only YAML emitter, modeled on <c>Utf8JsonWriter</c>. Calls describe the document as events (start and end
/// of mappings and sequences, keys, scalars, aliases, anchors, tags, comments and document markers) and the writer lays
/// them out as block or flow YAML.
/// </summary>
/// <remarks>
/// <para>Output is deterministic and always valid YAML for the calls made. Strings written with the default style are
/// plain when no YAML 1.2 core or YAML 1.1 reader could take them for anything but a string; otherwise they are quoted,
/// so <c>yes</c>, <c>on</c>, <c>012</c> or <c>1:20</c> stay strings for PyYAML and other YAML 1.1 readers.
/// Multi-line strings become literal block scalars in block context. Empty collections are written as <c>[]</c> and
/// <c>{}</c>.</para>
/// <para>Writing allocates nothing per call beyond the pooled output buffer of the stream overload. A writer is not
/// thread-safe. <see cref="Dispose"/> ends the last line and flushes; it never disposes the caller's stream or text
/// writer.</para>
/// </remarks>
public sealed class YamlWriter : IDisposable {
    private const int StreamBufferSize = 16384;
    private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

    private readonly TextWriter _text;
    private readonly bool _ownsText;
    private readonly YamlEmitter _emitter;
    private readonly YamlNodeBuilder? _builder;
    // Scratch for formatting a number before it is copied out, shared by the writers of a thread: nothing runs between
    // formatting into it and copying from it.
    [ThreadStatic] private static char[]? t_number;
    private bool _disposed;

    /// <summary>Creates a writer that encodes UTF-8 without a byte-order mark into a stream.</summary>
    /// <param name="utf8Stream">A writable stream. It is not disposed by the writer.</param>
    /// <param name="options">Layout settings, or <see langword="null"/> for the defaults.</param>
    public YamlWriter(Stream utf8Stream, YamlWriterOptions? options = null) {
        if (utf8Stream is null) throw new ArgumentNullException(nameof(utf8Stream));
        if (!utf8Stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(utf8Stream));
        _text = new PooledUtf8TextWriter(utf8Stream, Utf8, StreamBufferSize, leaveOpen: true);
        _ownsText = true;
        _emitter = new YamlEmitter(_text, options ?? YamlWriterOptions.Default);
    }

    /// <summary>Creates a writer over a caller-owned <see cref="TextWriter"/>.</summary>
    /// <param name="writer">The destination. It is not disposed by the writer.</param>
    /// <param name="options">Layout settings, or <see langword="null"/> for the defaults.</param>
    public YamlWriter(TextWriter writer, YamlWriterOptions? options = null) {
        _text = writer ?? throw new ArgumentNullException(nameof(writer));
        _emitter = new YamlEmitter(_text, options ?? YamlWriterOptions.Default);
    }

    /// <summary>Creates a writer that builds a detached node tree instead of text, for the serializer's node output.</summary>
    internal YamlWriter(YamlNodeBuilder builder) {
        _builder = builder;
        _text = TextWriter.Null;
        _emitter = new YamlEmitter(_text, YamlWriterOptions.Default);
    }

    /// <summary>Starts a document with an explicit <c>---</c> marker. The first document needs no call: writing its root starts it.</summary>
    /// <exception cref="InvalidOperationException">A collection is open, or the current document has no root yet.</exception>
    public void WriteDocumentStart() {
        CheckDisposed();
        if (_builder is not null) return;
        _emitter.StartDocument(explicitMarker: true);
    }

    /// <summary>Ends the current document with an explicit <c>...</c> marker. Starting the next document ends it implicitly.</summary>
    /// <exception cref="InvalidOperationException">No complete document is open.</exception>
    public void WriteDocumentEnd() {
        CheckDisposed();
        if (_builder is not null) return;
        _emitter.EndDocument();
    }

    /// <summary>Starts a mapping. Its entries follow as alternating keys and values.</summary>
    /// <param name="style"><see cref="YamlStyle.Block"/> (the default) or <see cref="YamlStyle.Flow"/>; inside a flow collection every collection is flow.</param>
    public void WriteStartMapping(YamlStyle style = YamlStyle.Block) {
        CheckDisposed();
        if (_builder is not null) _builder.StartCollection(mapping: true, CollectionStyle(style));
        else _emitter.StartCollection(mapping: true, CollectionStyle(style));
    }

    /// <summary>Ends the innermost mapping.</summary>
    /// <exception cref="InvalidOperationException">The innermost open collection is not a mapping, or its last key has no value.</exception>
    public void WriteEndMapping() {
        CheckDisposed();
        if (_builder is not null) _builder.EndCollection(mapping: true);
        else _emitter.EndCollection(mapping: true);
    }

    /// <summary>Starts a sequence.</summary>
    /// <param name="style"><see cref="YamlStyle.Block"/> (the default) or <see cref="YamlStyle.Flow"/>; inside a flow collection every collection is flow.</param>
    public void WriteStartSequence(YamlStyle style = YamlStyle.Block) {
        CheckDisposed();
        if (_builder is not null) _builder.StartCollection(mapping: false, CollectionStyle(style));
        else _emitter.StartCollection(mapping: false, CollectionStyle(style));
    }

    /// <summary>Ends the innermost sequence.</summary>
    /// <exception cref="InvalidOperationException">The innermost open collection is not a sequence.</exception>
    public void WriteEndSequence() {
        CheckDisposed();
        if (_builder is not null) _builder.EndCollection(mapping: false);
        else _emitter.EndCollection(mapping: false);
    }

    /// <summary>Writes a string mapping key with the default style. Complex keys are written with the node methods at a key position.</summary>
    /// <exception cref="InvalidOperationException">The writer is not at a key position.</exception>
    public void WriteKey(string key) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        CheckDisposed();
        if (_builder is not null) {
            if (!_builder.AtKey) throw new InvalidOperationException("WriteKey is valid only where a mapping expects a key.");
            _builder.Scalar(new YamlScalar(key));
            return;
        }
        if (!_emitter.AtKey) throw new InvalidOperationException("WriteKey is valid only where a mapping expects a key.");
        WriteChosen(key, _emitter.InFlow ? _emitter.ChooseStyle(key, YamlStyle.None) : YamlKeyStyles.Block(key), YamlStyle.None);
    }

    /// <summary>
    /// Writes a string mapping key as <see cref="WriteKey(string)"/> does, with its block-mapping style chosen in advance
    /// by <see cref="YamlKeyStyles.Block"/>; in a flow mapping the style is chosen again.
    /// </summary>
    internal void WriteKey(string key, YamlStyle blockStyle) {
        CheckDisposed();
        if (_builder is not null) {
            if (!_builder.AtKey) throw new InvalidOperationException("WriteKey is valid only where a mapping expects a key.");
            _builder.Scalar(new YamlScalar(key));
            return;
        }
        if (!_emitter.AtKey) throw new InvalidOperationException("WriteKey is valid only where a mapping expects a key.");
        WriteChosen(key, _emitter.InFlow ? _emitter.ChooseStyle(key, YamlStyle.None) : blockStyle, YamlStyle.None);
    }

    /// <summary>
    /// Writes a string scalar, or a null scalar when <paramref name="value"/> is <see langword="null"/>. The writer uses
    /// <paramref name="style"/> when it can represent the value at this position and otherwise the most conservative
    /// style that can; <see cref="YamlStyle.None"/> (the default) chooses as described on <see cref="YamlWriter"/>.
    /// A plain style is honored even when the text reads as another type, so a caller can write <c>on</c> on purpose.
    /// </summary>
    public void WriteScalar(string? value, YamlStyle style = YamlStyle.None) {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(value is null ? new YamlScalar((string?)null) : style == YamlStyle.None ? new YamlScalar(value) : new YamlScalar(value, ScalarStyle(style)));
            return;
        }
        if (value is null) {
            _emitter.WritePlain("null".AsSpan());
            return;
        }
        WriteChosen(value, _emitter.ChooseStyle(value, ScalarStyle(style)), style);
    }

    /// <summary>
    /// Writes a scalar in its chosen style. A default-style choice of plain or single quotes guarantees one line, so
    /// the emitter skips its line-break scans.
    /// </summary>
    private void WriteChosen(string value, YamlStyle chosen, YamlStyle requested) {
        if (requested == YamlStyle.None && chosen is YamlStyle.Plain or YamlStyle.SingleQuoted && value.Length <= YamlScalarFormat.MaxImplicitKeyLength) {
            _emitter.WriteSingleLine(value, chosen);
        } else {
            _emitter.WriteScalar(value, chosen);
        }
    }

    /// <summary>Writes an integer as a plain scalar.</summary>
    public void WriteScalar(long value) {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(new YamlScalar(value));
            return;
        }
        char[] number = t_number ??= new char[48];
        int length = YamlScalarFormat.Format(value, number);
        _emitter.WritePlain(number.AsSpan(0, length));
    }

    /// <summary>
    /// Writes a floating-point number as a plain scalar that YAML 1.2 and YAML 1.1 readers both read as the same float:
    /// <c>1.0</c>, <c>2.5e-07</c>, <c>.inf</c>, <c>-.inf</c> or <c>.nan</c>.
    /// </summary>
    public void WriteScalar(double value) {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(new YamlScalar(value));
            return;
        }
#if NET8_0_OR_GREATER
        char[] number = t_number ??= new char[48];
        int length = YamlScalarFormat.Format(value, number);
        _emitter.WritePlain(number.AsSpan(0, length));
#else
        _emitter.WritePlain(YamlScalarFormat.Format(value).AsSpan());
#endif
    }

    /// <summary>Writes <c>true</c> or <c>false</c>.</summary>
    public void WriteScalar(bool value) {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(new YamlScalar(value));
            return;
        }
        _emitter.WritePlain(YamlScalarFormat.Format(value).AsSpan());
    }

    /// <summary>Writes a null scalar (<c>null</c>).</summary>
    internal void WriteNull() {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(new YamlScalar((string?)null));
            return;
        }
        _emitter.WritePlain("null".AsSpan());
    }

    /// <summary>Writes a mapping key known to be safe as a plain scalar in every context, skipping the style choice.</summary>
    internal void WritePlainKey(string key) {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(new YamlScalar(key, YamlStyle.Plain));
            return;
        }
        _emitter.WriteSingleLine(key, YamlStyle.Plain);
    }

    /// <summary>
    /// Writes a block mapping entry whose key is safe as a plain scalar everywhere (or, with <paramref name="checkKey"/>,
    /// whose key is checked for the default style) and whose string value gets its default style, as
    /// <see cref="WritePlainKey"/> or <see cref="WriteKey(string)"/> and <see cref="WriteScalar(string?, YamlStyle)"/> would.
    /// A plain value goes out with the key in one step; the style of any other is chosen once, not again by the scalar
    /// path. Returns false, having written nothing, when the position or the key needs the general path.
    /// </summary>
    internal bool TryWriteStringEntry(string key, string value, bool checkKey = false) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockKey) return false;
        if (checkKey && !IsPlainKey(key)) return false;
        ReadOnlySpan<char> text = value.AsSpan();
        if (YamlScalarFormat.CanBePlain(text, flow: false) && !YamlScalarFormat.IsSchemaAmbiguous(text)) {
            _emitter.WriteEntry(key, text);
            return true;
        }
        _emitter.WriteSingleLine(key, YamlStyle.Plain);
        WriteChosen(value, YamlScalarFormat.ChooseNotPlain(text, flow: false, key: false, YamlStyle.SingleQuoted), YamlStyle.None);
        return true;
    }

    /// <summary>Whether a block mapping key gets the default style plain on one line, as an implicit key.</summary>
    internal static bool IsPlainKey(string key) => YamlKeyStyles.IsPlain(key);

    /// <summary>
    /// Writes a block sequence item for a string in its default style, as <see cref="WriteScalar(string?, YamlStyle)"/>
    /// would; a plain item goes out in one step. Returns false, having written nothing, when the position needs the
    /// general path.
    /// </summary>
    internal bool TryWriteStringItem(string value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockItem) return false;
        ReadOnlySpan<char> text = value.AsSpan();
        if (YamlScalarFormat.CanBePlain(text, flow: false) && !YamlScalarFormat.IsSchemaAmbiguous(text)) {
            _emitter.WriteItem(text);
            return true;
        }
        WriteChosen(value, YamlScalarFormat.ChooseNotPlain(text, flow: false, key: false, YamlStyle.SingleQuoted), YamlStyle.None);
        return true;
    }

    /// <summary>Writes a block mapping entry with a plain-safe key and an integer value in one step; see <see cref="TryWriteStringEntry"/>.</summary>
    internal bool TryWritePlainEntry(string key, long value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockKey) return false;
        char[] number = t_number ??= new char[48];
        int length = YamlScalarFormat.Format(value, number);
        _emitter.WriteEntry(key, number.AsSpan(0, length));
        return true;
    }

    /// <summary>Writes a block mapping entry with a plain-safe key and a floating-point value in one step; see <see cref="TryWriteStringEntry"/>.</summary>
    internal bool TryWritePlainEntry(string key, double value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockKey) return false;
#if NET8_0_OR_GREATER
        char[] number = t_number ??= new char[48];
        int length = YamlScalarFormat.Format(value, number);
        _emitter.WriteEntry(key, number.AsSpan(0, length));
#else
        _emitter.WriteEntry(key, YamlScalarFormat.Format(value).AsSpan());
#endif
        return true;
    }

    /// <summary>
    /// Writes a block mapping entry with a plain-safe key and a decimal value in its invariant spelling, as the decimal
    /// converter writes it, in one step; see <see cref="TryWriteStringEntry"/>.
    /// </summary>
    internal bool TryWritePlainEntry(string key, decimal value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockKey) return false;
#if NET8_0_OR_GREATER
        // A decimal has at most 29 digits, a sign and a point, which the number buffer holds.
        char[] number = t_number ??= new char[48];
        int length = YamlScalarFormat.Format(value, number);
        _emitter.WriteEntry(key, number.AsSpan(0, length));
#else
        _emitter.WriteEntry(key, value.ToString(CultureInfo.InvariantCulture).AsSpan());
#endif
        return true;
    }

    /// <summary>Writes a block mapping entry with a plain-safe key and a boolean value in one step; see <see cref="TryWriteStringEntry"/>.</summary>
    internal bool TryWritePlainEntry(string key, bool value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockKey) return false;
        _emitter.WriteEntry(key, YamlScalarFormat.Format(value).AsSpan());
        return true;
    }

    /// <summary>Writes a block mapping entry with a plain-safe key and a null value in one step; see <see cref="TryWriteStringEntry"/>.</summary>
    internal bool TryWriteNullEntry(string key) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockKey) return false;
        _emitter.WriteEntry(key, "null".AsSpan());
        return true;
    }

    /// <summary>
    /// Writes a block sequence item for an integer, a floating-point number, a boolean or a null, as the scalar methods
    /// would, in one step. Returns false, having written nothing, when the position needs the general path.
    /// </summary>
    internal bool TryWritePlainItem(long value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockItem) return false;
        char[] number = t_number ??= new char[48];
        _emitter.WriteItem(number.AsSpan(0, YamlScalarFormat.Format(value, number)));
        return true;
    }

    /// <inheritdoc cref="TryWritePlainItem(long)"/>
    internal bool TryWritePlainItem(double value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockItem) return false;
#if NET8_0_OR_GREATER
        char[] number = t_number ??= new char[48];
        _emitter.WriteItem(number.AsSpan(0, YamlScalarFormat.Format(value, number)));
#else
        _emitter.WriteItem(YamlScalarFormat.Format(value).AsSpan());
#endif
        return true;
    }

    /// <inheritdoc cref="TryWritePlainItem(long)"/>
    internal bool TryWritePlainItem(bool value) {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockItem) return false;
        _emitter.WriteItem(YamlScalarFormat.Format(value).AsSpan());
        return true;
    }

    /// <inheritdoc cref="TryWritePlainItem(long)"/>
    internal bool TryWriteNullItem() {
        CheckDisposed();
        if (_builder is not null || !_emitter.AtBlockItem) return false;
        _emitter.WriteItem("null".AsSpan());
        return true;
    }

    /// <summary>Writes text already in a canonical plain spelling, such as a formatted number or timestamp.</summary>
    internal void WritePlainValue(ReadOnlySpan<char> value) {
        CheckDisposed();
        if (_builder is not null) {
            _builder.Scalar(new YamlScalar(value.ToString(), YamlStyle.Plain));
            return;
        }
        _emitter.WritePlain(value);
    }

    /// <summary>Writes an alias (<c>*name</c>) to an anchor written earlier in the document.</summary>
    /// <exception cref="ArgumentException">The name is not a valid anchor name.</exception>
    public void WriteAlias(string anchor) {
        if (anchor is null) throw new ArgumentNullException(nameof(anchor));
        CheckDisposed();
        if (_builder is not null) _builder.Alias(anchor);
        else _emitter.WriteAlias(anchor);
    }

    /// <summary>Sets the anchor (<c>&amp;name</c>) of the next node.</summary>
    /// <exception cref="ArgumentException">The name is empty or contains spaces or <c>, [ ] { }</c>.</exception>
    public void WriteAnchor(string name) {
        if (name is null) throw new ArgumentNullException(nameof(name));
        CheckDisposed();
        if (_builder is not null) _builder.SetAnchor(name);
        else _emitter.SetAnchor(name);
    }

    /// <summary>
    /// Sets the tag of the next node, in resolved form as <see cref="YamlNode.Tag"/> reports it:
    /// <c>tag:yaml.org,2002:str</c> is written <c>!!str</c>, local tags such as <c>!Ref</c> as they are, and other URIs verbatim.
    /// </summary>
    public void WriteTag(string tag) {
        if (tag is null) throw new ArgumentNullException(nameof(tag));
        CheckDisposed();
        if (_builder is not null) _builder.SetTag(tag);
        else _emitter.SetTag(tag);
    }

    /// <summary>
    /// Writes comment lines (<c># text</c>) at the indentation of the next entry; each line of <paramref name="text"/>
    /// becomes one comment. Comments cannot be written between a key and its value or inside a flow collection.
    /// </summary>
    public void WriteComment(string text) {
        if (text is null) throw new ArgumentNullException(nameof(text));
        CheckDisposed();
        if (_builder is not null) return;
        _emitter.WriteCommentLines(text);
    }

    /// <summary>Flushes written text to the destination without ending the current line.</summary>
    public void Flush() {
        CheckDisposed();
        _text.Flush();
    }

    /// <summary>Ends the last line, flushes and releases the pooled buffer. The destination stays open.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        try {
            // Ending the last line is best effort: disposing a writer with open collections must not throw.
            if (_emitter.Documents > 0 && _emitter.IsComplete) _emitter.Finish();
        } finally {
            _emitter.Release();
            if (_ownsText) _text.Dispose();
            else _text.Flush();
        }
    }

    private static YamlStyle CollectionStyle(YamlStyle style) => style switch {
        YamlStyle.None or YamlStyle.Block => YamlStyle.Block,
        YamlStyle.Flow => YamlStyle.Flow,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "A collection style is Block or Flow.")
    };

    private static YamlStyle ScalarStyle(YamlStyle style) => style switch {
        YamlStyle.None or YamlStyle.Plain or YamlStyle.SingleQuoted or YamlStyle.DoubleQuoted or YamlStyle.Literal or YamlStyle.Folded => style,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "A scalar style is None, Plain, SingleQuoted, DoubleQuoted, Literal or Folded.")
    };

    private void CheckDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(YamlWriter));
    }
}
