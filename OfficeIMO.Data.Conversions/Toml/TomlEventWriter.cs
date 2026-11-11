using OfficeIMO.Toml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Writes data as TOML through <see cref="TomlWriter"/>. TOML writes a table's key/value pairs before its sub-tables, so
/// the sink holds the document and writes it when the source ends. The held document is a flat list of node records and
/// one character buffer, both pooled, rather than an object per value, so a large array of tables costs no garbage
/// collection work while it is held. Objects become <c>[table]</c> sections and arrays whose items are all objects become
/// <c>[[array]]</c> sections; other values stay on their key's line, with objects inside arrays as inline tables.
/// Everything TOML cannot hold is reported: nulls are omitted, integers beyond 64 bits become strings, a root that is not
/// an object is wrapped in <see cref="DataConversionOptions.RootName"/>, and members that precede sub-tables move after
/// their table's values.
/// </summary>
internal sealed class TomlEventWriter : DataEventWriter {
    private const int None = -1;
    private readonly TomlWriter _toml;
    private readonly List<int> _documents = new();
    private readonly Dictionary<string, int> _reported = new(StringComparer.Ordinal);

    // The held document: node records linked first-child/next-sibling, scalar text in one character buffer, and member
    // keys as indexes into a table of distinct key strings.
    private Node[] _nodes = ArrayPool<Node>.Shared.Rent(1024);
    private int _nodeCount;
    private char[] _chars = ArrayPool<char>.Shared.Rent(16 * 1024);
    private int _charCount;
    private readonly List<string> _keys = new();
    private readonly int[] _keySlots = CreateKeySlots();

    // Open containers, innermost last, with the last child attached to each.
    private int[] _open = new int[16];
    private int[] _openLast = new int[16];
    private int _depth;

    private int _pendingKey = None;
    private List<string>? _pendingKeyComments;
    private List<string>? _pendingComments;
    // Comments above a member (a node with a key) or an item or document, and comments after a container's last child,
    // by node.
    private Dictionary<int, List<string>>? _comments;
    private Dictionary<int, List<string>>? _endComments;
    private bool _complete;

    private enum NodeType : byte {
        Object,
        Array,
        Scalar
    }

    private struct Node {
        public NodeType Type;
        public DataScalarKind Kind;
        /// <summary>For an array: every item is an object.</summary>
        public bool AllObjects;
        /// <summary>For a scalar: <see cref="Integer"/> holds the value (a boolean as 0 or 1, or a 64-bit integer).</summary>
        public bool Held;
        /// <summary>The member key in the key table, or <see cref="None"/> for an array item or a document.</summary>
        public int Key;
        public int First;
        public int Next;
        public int Count;
        public int TextStart;
        public int TextLength;
        public long Integer;
    }

    public TomlEventWriter(Stream output, DataConversionContext context) : base(context) {
        _toml = new TomlWriter(output, context.Options.TomlWriter, leaveOpen: true);
    }

    private static int[] CreateKeySlots() {
        var slots = new int[256];
        slots.AsSpan().Fill(None);
        return slots;
    }

    public override void WriteStartDocument() {
    }

    public override void WriteEndDocument() {
    }

    public override void WriteStartObject() => Open(NodeType.Object);

    public override void WriteStartArray() => Open(NodeType.Array);

    public override void WriteEndObject() => Close();

    public override void WriteEndArray() => Close();

    public override void WritePropertyName(ReadOnlySpan<char> name) {
        _pendingKey = KeyIndex(name);
        _pendingKeyComments = _pendingComments;
        _pendingComments = null;
    }

    // Rows of records repeat the same keys, so one string per distinct key is shared across the held document.
    private int KeyIndex(ReadOnlySpan<char> name) {
        uint hash = 2166136261;
        foreach (char c in name) hash = (hash ^ c) * 16777619;
        ref int slot = ref _keySlots[hash & 255];
        if (slot != None && name.SequenceEqual(_keys[slot].AsSpan())) return slot;
        _keys.Add(name.ToString());
        return slot = _keys.Count - 1;
    }

    private int KeyIndex(string name) {
        _keys.Add(name);
        return _keys.Count - 1;
    }

    // Booleans and 64-bit integers are held as values; other scalars keep their text.
    public override void WriteScalar(DataScalarKind kind, ReadOnlySpan<char> text) {
        int index = NewNode(NodeType.Scalar);
        ref Node node = ref _nodes[index];
        node.Kind = kind;
        if (kind == DataScalarKind.Boolean) {
            node.Held = true;
            node.Integer = text.Length == 4 ? 1 : 0;
        } else if (kind == DataScalarKind.Integer && TryParseInteger(text, out long integer)) {
            node.Held = true;
            node.Integer = integer;
        } else if (kind != DataScalarKind.Null) {
            if (_chars.Length - _charCount < text.Length) Grow(ref _chars, _charCount + text.Length);
            text.CopyTo(new Span<char>(_chars, _charCount, text.Length));
            node.TextStart = _charCount;
            node.TextLength = text.Length;
            _charCount += text.Length;
        }

        Attach(index);
    }

    private static bool TryParseInteger(ReadOnlySpan<char> text, out long value) {
        value = 0;
        bool negative = text.Length > 0 && text[0] == '-';
        int start = negative ? 1 : 0;
        if (text.Length == start || text.Length - start > 18) return false;
        for (int i = start; i < text.Length; i++) {
            char c = text[i];
            if (c < '0' || c > '9') return false;
            value = value * 10 + (c - '0');
        }

        if (negative) value = -value;
        return true;
    }

    public override void WriteComment(ReadOnlySpan<char> text) => (_pendingComments ??= new List<string>()).Add(text.ToString());

    private int NewNode(NodeType type) {
        if (_nodeCount == _nodes.Length) Grow(ref _nodes, _nodeCount + 1);
        int index = _nodeCount++;
        _nodes[index] = new Node { Type = type, AllObjects = true, Key = None, First = None, Next = None };
        return index;
    }

    private static void Grow<T>(ref T[] array, int minimum) {
        T[] larger = ArrayPool<T>.Shared.Rent(Math.Max(minimum, array.Length * 2));
        Array.Copy(array, larger, array.Length);
        ArrayPool<T>.Shared.Return(array);
        array = larger;
    }

    private void Open(NodeType type) {
        int index = NewNode(type);
        Attach(index);
        if (_depth == _open.Length) {
            Array.Resize(ref _open, _depth * 2);
            Array.Resize(ref _openLast, _depth * 2);
        }

        _open[_depth] = index;
        _openLast[_depth] = None;
        _depth++;
    }

    private void Close() {
        int index = _open[--_depth];
        if (_pendingComments != null) {
            (_endComments ??= new Dictionary<int, List<string>>())[index] = _pendingComments;
            _pendingComments = null;
        }
    }

    private void Attach(int index) {
        if (_depth == 0) {
            SetComments(index, _pendingComments);
            _pendingComments = null;
            _documents.Add(index);
            return;
        }

        int parent = _open[_depth - 1];
        ref Node container = ref _nodes[parent];
        if (container.Type == NodeType.Object) {
            _nodes[index].Key = _pendingKey == None ? KeyIndex(string.Empty) : _pendingKey;
            SetComments(index, _pendingKeyComments);
            _pendingKey = None;
            _pendingKeyComments = null;
        } else {
            SetComments(index, _pendingComments);
            _pendingComments = null;
            if (_nodes[index].Type != NodeType.Object) container.AllObjects = false;
        }

        int last = _openLast[_depth - 1];
        if (last == None) container.First = index;
        else _nodes[last].Next = index;
        _openLast[_depth - 1] = index;
        container.Count++;
    }

    private void SetComments(int index, List<string>? comments) {
        if (comments != null) (_comments ??= new Dictionary<int, List<string>>())[index] = comments;
    }

    private List<string>? CommentsOf(int index) => _comments != null && _comments.TryGetValue(index, out List<string>? list) ? list : null;

    private List<string>? EndCommentsOf(int index) => _endComments != null && _endComments.TryGetValue(index, out List<string>? list) ? list : null;

    private bool IsNull(int index) => _nodes[index].Type == NodeType.Scalar && _nodes[index].Kind == DataScalarKind.Null;

    /// <summary>Whether the node is written as <c>[[name]]</c> sections: a non-empty array of objects only.</summary>
    private bool IsTableArray(int index) => _nodes[index].Type == NodeType.Array && _nodes[index].Count > 0 && _nodes[index].AllObjects;

    private bool IsSection(int index) => _nodes[index].Type == NodeType.Object || IsTableArray(index);

    private string Text(int index) => new string(_chars, _nodes[index].TextStart, _nodes[index].TextLength);

    public override void Complete() {
        if (_complete) return;
        _complete = true;
        Context.ItemsWritten = 1;
        int root;
        if (_documents.Count == 0) {
            root = NewNode(NodeType.Object);
        } else if (_documents.Count == 1 && _nodes[_documents[0]].Type == NodeType.Object) {
            root = _documents[0];
        } else {
            Context.Report(OfficeDiagnosticCodes.RootWrapped, OfficeConversionLossKind.Approximation,
                "TOML documents are tables, so the data was wrapped in the key '" + Context.Options.RootName + "' (RootName).", null);
            int value;
            if (_documents.Count == 1) {
                value = _documents[0];
            } else {
                // The documents become the items of one array; their own comments stay with them.
                value = NewNode(NodeType.Array);
                int previous = None;
                foreach (int document in _documents) {
                    if (previous == None) _nodes[value].First = document;
                    else _nodes[previous].Next = document;
                    previous = document;
                    if (_nodes[document].Type != NodeType.Object) _nodes[value].AllObjects = false;
                    _nodes[value].Count++;
                }
            }

            _nodes[value].Key = KeyIndex(Context.Options.RootName);
            root = NewNode(NodeType.Object);
            _nodes[root].First = value;
            _nodes[root].Count = 1;
        }

        if (_pendingComments != null) {
            List<string>? end = EndCommentsOf(root);
            (_endComments ??= new Dictionary<int, List<string>>())[root] = end == null ? _pendingComments : end.Concat(_pendingComments).ToList();
            _pendingComments = null;
        }

        // A wrapped document's comments sit above its member; only a document that is the root table writes them first.
        if (root == (_documents.Count > 0 ? _documents[0] : None)) WriteComments(CommentsOf(root));
        WriteTable(root, new List<string>());
        _toml.Flush();
    }

    // ------------------------------------------------------------------ sections

    private void WriteTable(int table, List<string> path) {
        bool sawSection = false;
        for (int member = _nodes[table].First; member != None; member = _nodes[member].Next) {
            if (IsSection(member)) {
                sawSection = true;
                continue;
            }

            if (sawSection) {
                Report(OfficeDiagnosticCodes.KeyOrderChanged, OfficeConversionLossKind.Approximation,
                    "TOML writes a table's values before its sub-tables, so values that followed a nested object or array of objects were moved before it.",
                    path, null);
                sawSection = false;
            }

            WriteComments(CommentsOf(member));
            string key = _keys[_nodes[member].Key];
            if (IsNull(member)) {
                ReportNull(path, key);
                continue;
            }

            _toml.WriteKey(key);
            WriteValue(member, path, key);
        }

        for (int member = _nodes[table].First; member != None; member = _nodes[member].Next) {
            if (!IsSection(member)) continue;
            // The path is shared down the tree: push the key, write the section, pop it.
            path.Add(_keys[_nodes[member].Key]);
            string[] keys = path.ToArray();
            List<string>? comments = CommentsOf(member);
            if (_nodes[member].Type == NodeType.Object) {
                WriteComments(comments);
                // A table that holds only sub-tables needs no header of its own: [a.b] defines a.
                if (HasValues(member) || _nodes[member].Count == 0 || comments != null) _toml.WriteTableHeader(keys);
                WriteTable(member, path);
            } else {
                bool first = true;
                for (int item = _nodes[member].First; item != None; item = _nodes[item].Next) {
                    WriteComments(first ? Merge(comments, CommentsOf(item)) : CommentsOf(item));
                    first = false;
                    _toml.WriteArrayOfTablesHeader(keys);
                    WriteTable(item, path);
                }

                WriteComments(EndCommentsOf(member));
            }

            path.RemoveAt(path.Count - 1);
        }

        WriteComments(EndCommentsOf(table));
    }

    private bool HasValues(int table) {
        for (int member = _nodes[table].First; member != None; member = _nodes[member].Next) {
            if (!IsSection(member)) return true;
        }

        return false;
    }

    private static List<string>? Merge(List<string>? first, List<string>? second) =>
        first == null ? second : second == null ? first : first.Concat(second).ToList();

    private void WriteComments(List<string>? comments) {
        if (comments == null) return;
        foreach (string comment in comments) {
            foreach (string line in comment.Split('\n')) _toml.WriteComment(Printable(line.TrimEnd('\r')));
        }
    }

    // TOML comments cannot hold control characters other than tab.
    private static string Printable(string text) {
        foreach (char c in text) {
            if (c < ' ' && c != '\t' || c == '\u007F') {
                var builder = new StringBuilder(text.Length);
                foreach (char ch in text) builder.Append(ch < ' ' && ch != '\t' || ch == '\u007F' ? ' ' : ch);
                return builder.ToString();
            }
        }

        return text;
    }

    // ------------------------------------------------------------------ values on a key's line

    private void WriteValue(int value, List<string> path, string key) {
        switch (_nodes[value].Type) {
            case NodeType.Scalar:
                WriteScalar(value, path, key);
                break;
            case NodeType.Array:
                InlineComments(value);
                _toml.WriteStartArray();
                for (int item = _nodes[value].First; item != None; item = _nodes[item].Next) {
                    InlineComments(item);
                    if (IsNull(item)) {
                        ReportNull(path, key);
                        continue;
                    }

                    WriteValue(item, path, key);
                }

                _toml.WriteEndArray();
                break;
            default:
                InlineComments(value);
                _toml.WriteStartInlineTable();
                for (int member = _nodes[value].First; member != None; member = _nodes[member].Next) {
                    if (CommentsOf(member) != null) OmitComments();
                    string memberKey = _keys[_nodes[member].Key];
                    if (IsNull(member)) {
                        ReportNull(path, key + "." + memberKey);
                        continue;
                    }

                    _toml.WriteKey(memberKey);
                    WriteValue(member, path, memberKey);
                }

                _toml.WriteEndInlineTable();
                break;
        }
    }

    // Comments above an array item, or after a value's last child, have no place in TOML 1.0 inside an array or inline
    // table. A member's own comments sit above its key and are written there.
    private void InlineComments(int node) {
        if ((_nodes[node].Key == None && CommentsOf(node) != null) || EndCommentsOf(node) != null) OmitComments();
    }

    private void OmitComments() =>
        Context.Report(OfficeDiagnosticCodes.CommentsOmitted, OfficeConversionLossKind.Omission,
            "Comments inside arrays and inline tables were dropped, because TOML 1.0 has no place for them.", null);

    private void WriteScalar(int value, List<string> path, string key) {
        Node node = _nodes[value];
        if (node.Held) {
            if (node.Kind == DataScalarKind.Boolean) _toml.WriteBoolean(node.Integer != 0);
            else _toml.WriteInteger(node.Integer);
            return;
        }

        string text = Text(value);
        switch (node.Kind) {
            case DataScalarKind.Integer:
                if (long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long integer)) {
                    _toml.WriteInteger(integer);
                } else {
                    OutOfRange(path, key, "Integers outside the 64-bit range TOML holds were written as strings.");
                    _toml.WriteString(text);
                }

                break;
            case DataScalarKind.Float:
                WriteFloat(text, path, key);
                break;
            case DataScalarKind.String:
                _toml.WriteString(text);
                break;
            default:
                if (TryParseDateTime(node.Kind, text, out TomlDateTime date)) {
                    _toml.WriteDateTime(date);
                } else {
                    Report(OfficeDiagnosticCodes.DatetimeAsString, OfficeConversionLossKind.Approximation,
                        "Date-times that are not in RFC 3339 form were written as strings.", path, key);
                    _toml.WriteString(text);
                }

                break;
        }
    }

    private void WriteFloat(string text, List<string> path, string key) {
        if (DataScalarKinds.IsNonFinite(text.AsSpan())) {
            _toml.WriteFloat(text[0] == 'N' ? double.NaN : text[0] == '-' ? double.NegativeInfinity : double.PositiveInfinity);
            return;
        }

        double number = double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (double.IsInfinity(number)) {
            OutOfRange(path, key, "Numbers beyond the range of a TOML float were written as strings.");
            _toml.WriteString(text);
            return;
        }

        if (SignificantDigits(text) > 15 && !SameDecimal(text, number.ToString("R", CultureInfo.InvariantCulture))) {
            Report(OfficeDiagnosticCodes.NumberPrecision, OfficeConversionLossKind.Approximation,
                "TOML floats are 64-bit; numbers with more significant digits were rounded.", path, key);
        }

        _toml.WriteFloat(number);
    }

    private void OutOfRange(List<string> path, string key, string message) =>
        Report(OfficeDiagnosticCodes.NumberOutOfRange, OfficeConversionLossKind.Approximation, message, path, key);

    private void ReportNull(List<string> path, string key) =>
        Report(OfficeDiagnosticCodes.NullDropped, OfficeConversionLossKind.Omission,
            "TOML has no null, so null values were omitted.", path, key);

    // The report keeps the first ten locations per code, so a path is built only for those.
    private void Report(string code, OfficeConversionLossKind loss, string message, List<string> path, string? key) {
        _reported.TryGetValue(code, out int count);
        _reported[code] = count + 1;
        Context.Report(code, loss, message, count < 10 ? Path(path, key) : null);
    }

    private static string Path(List<string> path, string? key = null) {
        var builder = new StringBuilder();
        foreach (string segment in path) DataPath.AppendKey(builder, segment);
        if (key != null) DataPath.AppendKey(builder, key);
        return builder.Length == 0 ? "$" : builder.ToString();
    }

    // ------------------------------------------------------------------ number and date-time text

    // Any decimal with at most 15 significant digits survives a round trip through a double.
    private static int SignificantDigits(string text) {
        int count = 0, zeros = 0;
        bool leading = true;
        foreach (char c in text) {
            if (c is 'e' or 'E') break;
            if (c < '0' || c > '9') continue;
            if (c == '0') {
                if (!leading) zeros++;
                continue;
            }

            leading = false;
            count += zeros + 1;
            zeros = 0;
        }

        return count;
    }

    /// <summary>Whether two JSON number spellings denote the same decimal value.</summary>
    internal static bool SameDecimal(string left, string right) {
        Normalize(left, out bool leftNegative, out string leftDigits, out long leftExponent);
        Normalize(right, out bool rightNegative, out string rightDigits, out long rightExponent);
        if (leftDigits.Length == 0 && rightDigits.Length == 0) return true;
        return leftNegative == rightNegative && leftExponent == rightExponent && string.Equals(leftDigits, rightDigits, StringComparison.Ordinal);
    }

    private static void Normalize(string text, out bool negative, out string digits, out long exponent) {
        int i = 0;
        negative = text.Length > 0 && text[0] == '-';
        if (negative || (text.Length > 0 && text[0] == '+')) i++;
        var builder = new StringBuilder(text.Length);
        long scale = 0;
        bool fraction = false;
        for (; i < text.Length; i++) {
            char c = text[i];
            if (c == '.') {
                fraction = true;
                continue;
            }

            if (c is 'e' or 'E') break;
            builder.Append(c);
            if (fraction) scale--;
        }

        exponent = scale;
        if (i < text.Length) exponent += long.Parse(text.Substring(i + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        int start = 0;
        while (start < builder.Length && builder[start] == '0') start++;
        int end = builder.Length;
        while (end > start && builder[end - 1] == '0') {
            end--;
            exponent++;
        }

        digits = builder.ToString(start, end - start);
    }

    /// <summary>
    /// Parses RFC 3339 text of the given kind with the TOML reader, which accepts exactly TOML's date-time syntax (a
    /// superset of RFC 3339) and keeps nanoseconds.
    /// </summary>
    internal static bool TryParseDateTime(DataScalarKind kind, string text, out TomlDateTime value) {
        value = default;
        TomlTokenType expected = kind switch {
            DataScalarKind.OffsetDateTime => TomlTokenType.OffsetDateTime,
            DataScalarKind.LocalDateTime => TomlTokenType.LocalDateTime,
            DataScalarKind.LocalDate => TomlTokenType.LocalDate,
            _ => TomlTokenType.LocalTime
        };
        try {
            var reader = new TomlReader("v = " + text);
            if (!reader.Read() || reader.TokenType != TomlTokenType.Key || !reader.Read() || reader.TokenType != expected) return false;
            value = reader.GetDateTime();
            return !reader.Read();
        } catch (TomlException) {
            return false;
        }
    }

    public override void Dispose() {
        _toml.Dispose();
        if (_nodes.Length > 0) {
            ArrayPool<Node>.Shared.Return(_nodes);
            _nodes = Array.Empty<Node>();
        }

        if (_chars.Length > 0) {
            ArrayPool<char>.Shared.Return(_chars);
            _chars = Array.Empty<char>();
        }
    }
}
