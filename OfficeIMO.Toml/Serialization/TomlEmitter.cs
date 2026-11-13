namespace OfficeIMO.Toml;

/// <summary>
/// The serializer's output: key/value pairs, <c>[table]</c> and <c>[[table]]</c> headers, arrays and inline tables,
/// formatted by <see cref="TomlFormatter"/> as <see cref="TomlWriter"/> writes them. A <c>[table]</c> header waits
/// until the table writes a pair or ends empty, so a table that holds only sub-tables gets no header of its own, as
/// TOML files are written by hand. One emitter is reused per thread.
/// </summary>
internal sealed class TomlEmitter {
    [ThreadStatic]
    private static TomlEmitter? _cached;

    private readonly StringBuilder _text = new StringBuilder(256);
    private readonly List<string> _path = new List<string>();
    private string _newLine = "\n";
    private bool _headerPending;
    // The open arrays and inline tables: how many items each has, and whether it is a table.
    private int[] _counts = new int[8];
    private bool[] _tables = new bool[8];
    private int _depth;

    public static TomlEmitter Rent(string newLine) {
        TomlEmitter emitter = _cached ?? new TomlEmitter();
        _cached = null;
        emitter._newLine = newLine;
        return emitter;
    }

    /// <summary>Returns the emitter for reuse on this thread; a very large buffer is dropped instead.</summary>
    public void Return() {
        _text.Clear();
        _path.Clear();
        _headerPending = false;
        _depth = 0;
        if (_text.Capacity <= 1 << 20) _cached = this;
    }

    public StringBuilder Text => _text;

    /// <summary>Whether the emitter is inside an array or inline table, where every value is written inline.</summary>
    public bool IsInline => _depth > 0;

    /// <summary>Writes a key and <c> = </c>; <paramref name="keyText"/> is already formatted, as property keys are.</summary>
    public void WriteKeyText(string keyText) {
        if (_depth == 0) {
            if (_headerPending) FlushHeader();
        } else {
            _text.Append(_counts[_depth - 1]++ > 0 ? ", " : " ");
        }
        _text.Append(keyText);
    }

    /// <summary>Writes a key that is not formatted yet, such as a dictionary key.</summary>
    public void WriteKey(string key) {
        if (_depth == 0) {
            if (_headerPending) FlushHeader();
        } else {
            _text.Append(_counts[_depth - 1]++ > 0 ? ", " : " ");
        }
        TomlFormatter.AppendKey(_text, key);
        _text.Append(" = ");
    }

    /// <summary>Formats a key for <see cref="WriteKeyText"/>.</summary>
    public static string FormatKey(string key) {
        var text = new StringBuilder(key.Length + 5);
        TomlFormatter.AppendKey(text, key);
        return text.Append(" = ").ToString();
    }

    /// <summary>Ends a key/value pair: a line break outside arrays and inline tables.</summary>
    public void EndValue() {
        if (_depth == 0) _text.Append(_newLine);
    }

    public void WriteString(string value) {
        BeforeValue();
        TomlFormatter.AppendString(_text, value, TomlFormatter.FitStyle(value, TomlStringStyle.Basic));
    }

    public void WriteInteger(long value) {
        BeforeValue();
        TomlFormatter.AppendInteger(_text, value, TomlIntegerBase.Decimal, upperHex: false);
    }

    public void WriteFloat(double value) {
        BeforeValue();
        TomlFormatter.AppendFloat(_text, value);
    }

    /// <summary>Writes a finite number from its invariant text, adding <c>.0</c> when the text is an integer.</summary>
    public void WriteFloatText(string text) {
        BeforeValue();
        _text.Append(text);
        if (text.IndexOf('.') < 0 && text.IndexOf('E') < 0 && text.IndexOf('e') < 0) _text.Append(".0");
    }

    public void WriteBoolean(bool value) {
        BeforeValue();
        _text.Append(value ? "true" : "false");
    }

    public void WriteDateTime(TomlDateTime value) {
        BeforeValue();
        TomlFormatter.AppendDateTime(_text, value, 'T');
    }

    public void BeginArray() {
        BeforeValue();
        _text.Append('[');
        Push(table: false);
    }

    public void EndArray() {
        _depth--;
        _text.Append(']');
    }

    public void BeginInlineTable() {
        BeforeValue();
        _text.Append('{');
        Push(table: true);
    }

    public void EndInlineTable() {
        _depth--;
        _text.Append(_counts[_depth] > 0 ? " }" : "}");
    }

    /// <summary>Starts a sub-table of the current table; its <c>[header]</c> is written when it writes a pair or ends empty.</summary>
    public void PushTable(string key) {
        _path.Add(key);
        _headerPending = true;
    }

    public void PopTable() {
        if (_headerPending) FlushHeader();
        _path.RemoveAt(_path.Count - 1);
    }

    /// <summary>Adds the key of an array of tables to the path; <see cref="BeginTableArrayItem"/> starts each table.</summary>
    public void PushKey(string key) {
        _path.Add(key);
        _headerPending = false;
    }

    public void PopKey() => _path.RemoveAt(_path.Count - 1);

    /// <summary>Writes the <c>[[header]]</c> that starts the next table of the array of tables at the current path.</summary>
    public void BeginTableArrayItem() {
        _headerPending = false;
        if (_text.Length > 0) _text.Append(_newLine);
        _text.Append("[[");
        TomlFormatter.AppendKeys(_text, _path);
        _text.Append("]]").Append(_newLine);
    }

    private void FlushHeader() {
        _headerPending = false;
        if (_text.Length > 0) _text.Append(_newLine);
        _text.Append('[');
        TomlFormatter.AppendKeys(_text, _path);
        _text.Append(']').Append(_newLine);
    }

    private void BeforeValue() {
        if (_depth > 0 && !_tables[_depth - 1] && _counts[_depth - 1]++ > 0) _text.Append(", ");
    }

    private void Push(bool table) {
        if (_depth == _counts.Length) {
            Array.Resize(ref _counts, _depth * 2);
            Array.Resize(ref _tables, _depth * 2);
        }
        _counts[_depth] = 0;
        _tables[_depth] = table;
        _depth++;
    }
}

/// <summary>Per-call writing state: null handling, cycle detection and the pending discriminator of a derived type.</summary>
internal sealed class TomlWriteState {
    private const int MaxDepth = 1000;
    private object?[] _stack = new object?[8];
    private int _depth;

    public TomlWriteState(TomlSerializerOptions options) {
        Options = options;
        FailOnNull = options.NullHandling == TomlNullHandling.Error;
    }

    public TomlSerializerOptions Options { get; }

    public bool FailOnNull { get; }

    /// <summary>Set by a <see cref="TomlInlineTableAttribute"/> dictionary property for its entries; the dictionary clears it.</summary>
    public bool InlineEntries { get; set; }

    /// <summary>The discriminator key text (formatted) and value the next object writes first, for a derived type.</summary>
    public string? DiscriminatorKeyText { get; set; }

    public string? DiscriminatorValue { get; set; }

    /// <summary>Enters a container value, failing when it is already being written (a cycle) or nests too deep.</summary>
    public void Enter(object value, Type type) {
        for (int i = 0; i < _depth; i++) {
            if (ReferenceEquals(_stack[i], value)) throw TomlSerializationErrors.Cycle(type);
        }
        if (_depth == MaxDepth) throw TomlSerializationErrors.Cycle(type);
        if (_depth == _stack.Length) Array.Resize(ref _stack, _depth * 2);
        _stack[_depth++] = value;
    }

    public void Exit() => _stack[--_depth] = null;
}

/// <summary>Per-call reading state.</summary>
internal sealed class TomlReadState {
    public TomlReadState(TomlSerializerOptions options) {
        Options = options;
    }

    public TomlSerializerOptions Options { get; }
}
