namespace OfficeIMO.Toml;

/// <summary>Settings for <see cref="TomlWriter"/>.</summary>
public sealed class TomlWriterOptions {
    private string _newLine = "\n";

    /// <summary>Gets or sets the line break: <c>"\n"</c> (the default) or <c>"\r\n"</c>.</summary>
    /// <exception cref="ArgumentException">The value is not <c>"\n"</c> or <c>"\r\n"</c>.</exception>
    public string NewLine {
        get => _newLine;
        set => _newLine = value is "\n" or "\r\n" ? value : throw new ArgumentException("TOML line breaks are \"\\n\" or \"\\r\\n\".", nameof(value));
    }
}

/// <summary>
/// A forward-only TOML emitter, modeled on <c>Utf8JsonWriter</c>. Calls write headers, keys, values, arrays, inline
/// tables and comments in order; the writer adds the <c>=</c>, separators and line breaks, and rejects calls TOML does
/// not allow at that point. Output is deterministic and readable by TOML 1.0 and 1.1 parsers.
/// </summary>
/// <example>
/// <code>
/// using var writer = new TomlWriter(stream);
/// writer.WriteKey("name");
/// writer.WriteString("officeimo");
/// writer.WriteTableHeader("tool", "ruff");
/// writer.WriteKey("line-length");
/// writer.WriteInteger(120);
/// </code>
/// </example>
public sealed class TomlWriter : IDisposable {
    private enum Context : byte {
        Document,
        Array,
        InlineTable
    }

    private readonly TextWriter _writer;
    private readonly bool _leaveOpen;
    private readonly string _newLine;
    private readonly StringBuilder _buffer = new StringBuilder();
    private readonly List<Context> _contexts = new List<Context>();
    private readonly List<int> _counts = new List<int>();
    private bool _keyPending;
    private bool _wroteAnything;

    /// <summary>Creates a writer over a text writer.</summary>
    /// <param name="writer">The destination.</param>
    /// <param name="options">Writer settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="leaveOpen">Whether disposing the writer leaves <paramref name="writer"/> open.</param>
    public TomlWriter(TextWriter writer, TomlWriterOptions? options = null, bool leaveOpen = false) {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
        _leaveOpen = leaveOpen;
        _newLine = (options ?? new TomlWriterOptions()).NewLine;
    }

    /// <summary>Creates a writer that writes UTF-8 without a byte-order mark to a stream.</summary>
    /// <param name="stream">A writable stream.</param>
    /// <param name="options">Writer settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="leaveOpen">Whether disposing the writer leaves <paramref name="stream"/> open.</param>
    public TomlWriter(Stream stream, TomlWriterOptions? options = null, bool leaveOpen = false)
        : this(new StreamWriter(stream ?? throw new ArgumentNullException(nameof(stream)), new UTF8Encoding(false), 4096, leaveOpen), options, leaveOpen: false) {
    }

    private Context Current => _contexts.Count == 0 ? Context.Document : _contexts[_contexts.Count - 1];

    /// <summary>Writes a <c>[table]</c> header line, preceded by a blank line unless it is the first output.</summary>
    /// <param name="keys">The table's keys from the root, such as <c>"tool", "ruff"</c>.</param>
    /// <exception cref="InvalidOperationException">A key is waiting for its value, or an array or inline table is open.</exception>
    public void WriteTableHeader(params string[] keys) => WriteHeader(keys, "[", "]");

    /// <summary>Writes a <c>[[table]]</c> header line that starts a new table of an array of tables.</summary>
    /// <param name="keys">The array's keys from the root, such as <c>"package"</c>.</param>
    /// <exception cref="InvalidOperationException">A key is waiting for its value, or an array or inline table is open.</exception>
    public void WriteArrayOfTablesHeader(params string[] keys) => WriteHeader(keys, "[[", "]]");

    private void WriteHeader(string[] keys, string open, string close) {
        RequireKeys(keys);
        if (_keyPending || _contexts.Count > 0) throw new InvalidOperationException("A header cannot be written while a key/value pair, array or inline table is unfinished.");
        if (_wroteAnything) _writer.Write(_newLine);
        _buffer.Clear().Append(open);
        TomlFormatter.AppendKeys(_buffer, keys);
        _buffer.Append(close).Append(_newLine);
        Flush(_buffer);
    }

    /// <summary>Writes a key, or the dotted key made of several keys; the next call writes its value.</summary>
    /// <param name="keys">One key, or the segments of a dotted key.</param>
    /// <exception cref="InvalidOperationException">A key is already waiting for its value, or the writer is inside an array.</exception>
    public void WriteKey(params string[] keys) {
        RequireKeys(keys);
        if (_keyPending) throw new InvalidOperationException("The previous key has no value yet.");
        Context context = Current;
        if (context == Context.Array) throw new InvalidOperationException("Array items have no keys.");
        _buffer.Clear();
        if (context == Context.InlineTable && Separate()) _buffer.Append(", ");
        TomlFormatter.AppendKeys(_buffer, keys);
        _buffer.Append(" = ");
        Flush(_buffer);
        _keyPending = true;
    }

    /// <summary>Writes a string value in the given style, or as a basic string when that style cannot hold it.</summary>
    public void WriteString(string value, TomlStringStyle style = TomlStringStyle.Basic) {
        if (value is null) throw new ArgumentNullException(nameof(value));
        BeginValue();
        TomlFormatter.AppendString(_buffer, value, TomlFormatter.FitStyle(value, style));
        EndValue();
    }

    /// <summary>Writes an integer value.</summary>
    public void WriteInteger(long value) {
        BeginValue();
        TomlFormatter.AppendInteger(_buffer, value, TomlIntegerBase.Decimal, upperHex: false);
        EndValue();
    }

    /// <summary>Writes a float value; NaN and infinities are written as <c>nan</c>, <c>inf</c> and <c>-inf</c>.</summary>
    public void WriteFloat(double value) {
        BeginValue();
        TomlFormatter.AppendFloat(_buffer, value);
        EndValue();
    }

    /// <summary>Writes a boolean value.</summary>
    public void WriteBoolean(bool value) {
        BeginValue();
        _buffer.Append(value ? "true" : "false");
        EndValue();
    }

    /// <summary>Writes a date, time or date-time value.</summary>
    public void WriteDateTime(TomlDateTime value) {
        BeginValue();
        TomlFormatter.AppendDateTime(_buffer, value, 'T');
        EndValue();
    }

    /// <summary>Starts an array value; write its items, then <see cref="WriteEndArray"/>.</summary>
    public void WriteStartArray() => Open(Context.Array, "[");

    /// <summary>Ends the innermost array.</summary>
    /// <exception cref="InvalidOperationException">The innermost open value is not an array.</exception>
    public void WriteEndArray() => Close(Context.Array, "]");

    /// <summary>Starts an inline table; write its keys and values, then <see cref="WriteEndInlineTable"/>.</summary>
    public void WriteStartInlineTable() => Open(Context.InlineTable, "{");

    /// <summary>Ends the innermost inline table.</summary>
    /// <exception cref="InvalidOperationException">The innermost open value is not an inline table, or a key has no value.</exception>
    public void WriteEndInlineTable() => Close(Context.InlineTable, "}");

    /// <summary>Writes a comment on its own line.</summary>
    /// <exception cref="ArgumentException">The comment contains a line break or a control character.</exception>
    /// <exception cref="InvalidOperationException">A key/value pair, array or inline table is unfinished.</exception>
    public void WriteComment(string comment) {
        if (comment is null) throw new ArgumentNullException(nameof(comment));
        foreach (char ch in comment) {
            if (TomlScalars.IsForbiddenControl(ch)) throw new ArgumentException("A comment cannot contain line breaks or control characters.", nameof(comment));
        }
        if (_keyPending || _contexts.Count > 0) throw new InvalidOperationException("A comment can only be written between key/value pairs.");
        _buffer.Clear().Append('#');
        if (comment.Length > 0 && comment[0] != ' ') _buffer.Append(' ');
        _buffer.Append(comment).Append(_newLine);
        Flush(_buffer);
    }

    /// <summary>Flushes buffered output to the destination.</summary>
    public void Flush() => _writer.Flush();

    /// <summary>Flushes and, unless the writer was created to leave it open, disposes the destination.</summary>
    public void Dispose() {
        _writer.Flush();
        if (!_leaveOpen) _writer.Dispose();
    }

    private void Open(Context context, string bracket) {
        BeginValue();
        _buffer.Append(bracket);
        Flush(_buffer);
        _contexts.Add(context);
        _counts.Add(0);
    }

    private void Close(Context context, string bracket) {
        if (Current != context || _contexts.Count == 0) throw new InvalidOperationException("There is no open " + (context == Context.Array ? "array" : "inline table") + " to end.");
        if (_keyPending) throw new InvalidOperationException("The last key has no value.");
        bool empty = _counts[_counts.Count - 1] == 0;
        _contexts.RemoveAt(_contexts.Count - 1);
        _counts.RemoveAt(_counts.Count - 1);
        _buffer.Clear();
        if (context == Context.InlineTable && !empty) _buffer.Append(' ');
        _buffer.Append(bracket);
        EndValue();
    }

    private void BeginValue() {
        Context context = Current;
        _buffer.Clear();
        if (context == Context.Array) {
            if (Separate()) _buffer.Append(", ");
            return;
        }
        if (!_keyPending) throw new InvalidOperationException("A value needs a key; call WriteKey first.");
        _keyPending = false;
    }

    /// <summary>Counts an element of the innermost array or inline table and returns whether it needs a separator.</summary>
    private bool Separate() {
        int index = _counts.Count - 1;
        int count = _counts[index]++;
        if (Current == Context.InlineTable && count == 0) _buffer.Append(' ');
        return count > 0;
    }

    private void EndValue() {
        if (_contexts.Count == 0) _buffer.Append(_newLine);
        Flush(_buffer);
    }

    private void Flush(StringBuilder text) {
        _writer.Write(text.ToString());
        _wroteAnything = true;
    }

    private static void RequireKeys(string[] keys) {
        if (keys is null || keys.Length == 0) throw new ArgumentException("At least one key is required.", nameof(keys));
        foreach (string key in keys) {
            if (key is null) throw new ArgumentNullException(nameof(keys), "Keys cannot be null.");
        }
    }
}
