using OfficeIMO.Toml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Writes data as TOML through <see cref="TomlWriter"/>. TOML writes a table's key/value pairs before its sub-tables, so
/// the sink holds the document and writes it when the source ends; TOML documents are configuration-sized. Objects become
/// <c>[table]</c> sections and arrays whose items are all objects become <c>[[array]]</c> sections; other values stay on
/// their key's line, with objects inside arrays as inline tables. Everything TOML cannot hold is reported: nulls are
/// omitted, integers beyond 64 bits become strings, a root that is not an object is wrapped in
/// <see cref="DataConversionOptions.RootName"/>, and members that precede sub-tables move after their table's values.
/// </summary>
internal sealed class TomlEventWriter : DataEventWriter {
    private readonly TomlWriter _toml;
    private readonly Stack<Node> _open = new();
    private readonly List<Node> _documents = new();
    private List<string>? _comments;
    private string? _pendingKey;
    private List<string>? _pendingKeyComments;
    private bool _complete;
    private readonly Dictionary<string, int> _reported = new(StringComparer.Ordinal);

    private enum NodeType : byte {
        Object,
        Array,
        Scalar
    }

    private sealed class Node {
        public Node(NodeType type) {
            Type = type;
        }

        public NodeType Type { get; }
        public DataScalarKind Kind;
        public string? Text;
        public List<Member>? Members;
        public List<Node>? Items;
        /// <summary>Comments above this node when it is an array item.</summary>
        public List<string>? Comments;
        /// <summary>Comments after the last member or item.</summary>
        public List<string>? EndComments;

        public bool IsNull => Type == NodeType.Scalar && Kind == DataScalarKind.Null;

        /// <summary>Whether the node is written as <c>[[name]]</c> sections: a non-empty array of objects only.</summary>
        public bool IsTableArray {
            get {
                if (Type != NodeType.Array || Items!.Count == 0) return false;
                foreach (Node item in Items) {
                    if (item.Type != NodeType.Object) return false;
                }

                return true;
            }
        }
    }

    private readonly struct Member {
        public Member(string key, Node value, List<string>? comments) {
            Key = key;
            Value = value;
            Comments = comments;
        }

        public string Key { get; }
        public Node Value { get; }
        public List<string>? Comments { get; }

        public bool IsSection => Value.Type == NodeType.Object || Value.IsTableArray;
    }

    public TomlEventWriter(Stream output, DataConversionContext context) : base(context) {
        _toml = new TomlWriter(output, context.Options.TomlWriter, leaveOpen: true);
    }

    public override void WriteStartDocument() {
    }

    public override void WriteEndDocument() {
    }

    public override void WriteStartObject() => Open(new Node(NodeType.Object) { Members = new List<Member>() });

    public override void WriteStartArray() => Open(new Node(NodeType.Array) { Items = new List<Node>() });

    public override void WriteEndObject() => Close();

    public override void WriteEndArray() => Close();

    public override void WritePropertyName(ReadOnlySpan<char> name) {
        _pendingKey = name.ToString();
        _pendingKeyComments = _comments;
        _comments = null;
    }

    public override void WriteScalar(DataScalarKind kind, ReadOnlySpan<char> text) =>
        Attach(new Node(NodeType.Scalar) { Kind = kind, Text = kind == DataScalarKind.Null ? null : text.ToString() });

    public override void WriteComment(ReadOnlySpan<char> text) => (_comments ??= new List<string>()).Add(text.ToString());

    private void Open(Node node) {
        Attach(node);
        _open.Push(node);
    }

    private void Close() {
        Node node = _open.Pop();
        if (_comments != null) {
            node.EndComments = _comments;
            _comments = null;
        }
    }

    private void Attach(Node node) {
        if (_open.Count == 0) {
            node.Comments = _comments;
            _comments = null;
            _documents.Add(node);
            return;
        }

        Node parent = _open.Peek();
        if (parent.Type == NodeType.Object) {
            parent.Members!.Add(new Member(_pendingKey ?? string.Empty, node, _pendingKeyComments));
            _pendingKey = null;
            _pendingKeyComments = null;
        } else {
            node.Comments = _comments;
            _comments = null;
            parent.Items!.Add(node);
        }
    }

    public override void Complete() {
        if (_complete) return;
        _complete = true;
        Context.ItemsWritten = 1;
        Node root;
        if (_documents.Count == 0) {
            root = new Node(NodeType.Object) { Members = new List<Member>() };
        } else if (_documents.Count == 1 && _documents[0].Type == NodeType.Object) {
            root = _documents[0];
        } else {
            Context.Report(OfficeDiagnosticCodes.RootWrapped, OfficeConversionLossKind.Approximation,
                "TOML documents are tables, so the data was wrapped in the key '" + Context.Options.RootName + "' (RootName).", null);
            Node value;
            if (_documents.Count == 1) {
                value = _documents[0];
            } else {
                value = new Node(NodeType.Array) { Items = _documents };
            }

            root = new Node(NodeType.Object) { Members = new List<Member> { new Member(Context.Options.RootName, value, value.Comments) } };
            if (_documents.Count == 1) value.Comments = null;
        }

        if (_comments != null) {
            root.EndComments = root.EndComments == null ? _comments : root.EndComments.Concat(_comments).ToList();
            _comments = null;
        }

        WriteComments(root.Comments);
        WriteTable(root, new List<string>());
        _toml.Flush();
    }

    // ------------------------------------------------------------------ sections

    private void WriteTable(Node table, List<string> path) {
        bool sawSection = false;
        foreach (Member member in table.Members!) {
            if (member.IsSection) {
                sawSection = true;
                continue;
            }

            if (sawSection) {
                Report(OfficeDiagnosticCodes.KeyOrderChanged, OfficeConversionLossKind.Approximation,
                    "TOML writes a table's values before its sub-tables, so values that followed a nested object or array of objects were moved before it.",
                    path, null);
                sawSection = false;
            }

            WriteComments(member.Comments);
            if (member.Value.IsNull) {
                ReportNull(path, member.Key);
                continue;
            }

            _toml.WriteKey(member.Key);
            WriteValue(member.Value, path, member.Key);
        }

        foreach (Member member in table.Members) {
            if (!member.IsSection) continue;
            var childPath = new List<string>(path) { member.Key };
            string[] keys = childPath.ToArray();
            if (member.Value.Type == NodeType.Object) {
                WriteComments(member.Comments);
                // A table that holds only sub-tables needs no header of its own: [a.b] defines a.
                if (HasValues(member.Value) || member.Value.Members!.Count == 0 || member.Comments != null) _toml.WriteTableHeader(keys);
                WriteTable(member.Value, childPath);
                continue;
            }

            bool first = true;
            foreach (Node item in member.Value.Items!) {
                WriteComments(first ? Merge(member.Comments, item.Comments) : item.Comments);
                first = false;
                _toml.WriteArrayOfTablesHeader(keys);
                WriteTable(item, childPath);
            }

            WriteComments(member.Value.EndComments);
        }

        WriteComments(table.EndComments);
    }

    private static bool HasValues(Node table) {
        foreach (Member member in table.Members!) {
            if (!member.IsSection) return true;
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

    private void WriteValue(Node value, List<string> path, string key) {
        switch (value.Type) {
            case NodeType.Scalar:
                WriteScalar(value, path, key);
                break;
            case NodeType.Array:
                InlineComments(value);
                _toml.WriteStartArray();
                foreach (Node item in value.Items!) {
                    InlineComments(item);
                    if (item.IsNull) {
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
                foreach (Member member in value.Members!) {
                    if (member.Comments != null) OmitComments();
                    if (member.Value.IsNull) {
                        ReportNull(path, key + "." + member.Key);
                        continue;
                    }

                    _toml.WriteKey(member.Key);
                    WriteValue(member.Value, path, member.Key);
                }

                _toml.WriteEndInlineTable();
                break;
        }
    }

    private void InlineComments(Node node) {
        if (node.Comments != null || node.EndComments != null) OmitComments();
    }

    private void OmitComments() =>
        Context.Report(OfficeDiagnosticCodes.CommentsOmitted, OfficeConversionLossKind.Omission,
            "Comments inside arrays and inline tables were dropped, because TOML 1.0 has no place for them.", null);

    private void WriteScalar(Node value, List<string> path, string key) {
        string text = value.Text!;
        switch (value.Kind) {
            case DataScalarKind.Boolean:
                _toml.WriteBoolean(text.Length == 4);
                break;
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
                if (TryParseDateTime(value.Kind, text, out TomlDateTime date)) {
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

    public override void Dispose() => _toml.Dispose();
}
