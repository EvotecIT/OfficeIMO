using OfficeIMO.Toml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Walks a loaded <see cref="TomlDocument"/> as data events, without recursion. Tables are objects, arrays of tables
/// and arrays are arrays, and values keep their TOML type: integers as canonical decimal text (<c>0xFF</c> is
/// <c>255</c>), floats in their source spelling without underscores, and the four date-time kinds as RFC 3339 text.
/// Comments above a key or header come before its property name; a comment after a value follows it.
/// </summary>
internal sealed class TomlModelEventReader : DataEventReader {
    private readonly TomlDocument _document;
    private readonly Queue<Pending> _pending = new();
    private readonly Stack<Frame> _frames = new();
    private TomlNode? _location;
    private int _comments;
    private int _state;

    private readonly struct Pending {
        public Pending(DataTokenType type, DataScalarKind kind = default, string? text = null) {
            Type = type;
            Kind = kind;
            Text = text;
        }

        public DataTokenType Type { get; }
        public DataScalarKind Kind { get; }
        public string? Text { get; }
    }

    private sealed class Frame {
        public Frame(TomlNode node, IEnumerator<KeyValuePair<string, TomlNode>>? members, System.Collections.IEnumerator? items) {
            Node = node;
            Members = members;
            Items = items;
        }

        public TomlNode Node { get; }
        public IEnumerator<KeyValuePair<string, TomlNode>>? Members { get; }
        public System.Collections.IEnumerator? Items { get; }
    }

    private readonly bool _owned;

    /// <param name="document">The document to walk.</param>
    /// <param name="context">The conversion context.</param>
    /// <param name="owned">Whether the reader disposes the document, as it does for documents it loaded from a stream.</param>
    public TomlModelEventReader(TomlDocument document, DataConversionContext context, bool owned = false) : base(context) {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _owned = owned;
    }

    public override void Dispose() {
        if (_owned) _document.Dispose();
    }

    public override string? Location => _location is { } node && node.Start.Line > 0
        ? "line " + node.Start.Line.ToString(CultureInfo.InvariantCulture) + ", column " + node.Start.Column.ToString(CultureInfo.InvariantCulture)
        : null;

    protected override bool ReadCore() {
        while (_pending.Count == 0) {
            if (!Advance()) return false;
        }

        Pending next = _pending.Dequeue();
        switch (next.Type) {
            case DataTokenType.PropertyName: SetPropertyName(next.Text.AsSpan()); break;
            case DataTokenType.Scalar: SetScalar(next.Kind, next.Text.AsSpan()); break;
            case DataTokenType.Comment: SetComment(next.Text.AsSpan()); break;
            default: Set(next.Type); break;
        }

        return true;
    }

    private bool Advance() {
        if (_state == 0) {
            _state = 1;
            _pending.Enqueue(new Pending(DataTokenType.StartDocument));
            Begin(_document.Root);
            return true;
        }

        if (_frames.Count == 0) {
            if (_state == 2) return false;
            _state = 2;
            _pending.Enqueue(new Pending(DataTokenType.EndDocument));
            return true;
        }

        Frame frame = _frames.Peek();
        if (frame.Members != null) {
            if (frame.Members.MoveNext()) {
                KeyValuePair<string, TomlNode> member = frame.Members.Current;
                TomlNode value = member.Value;
                LeadingComments(value);
                // A header's own trailing comment describes the table, so it travels before the table's members.
                if (value is TomlTable { TableKind: TomlTableKind.Header } or TomlTableArray) TrailingComment(value);
                _pending.Enqueue(new Pending(DataTokenType.PropertyName, text: member.Key));
                Begin(value);
                return true;
            }

            _frames.Pop();
            if (_frames.Count == 0) ReportUnattachedComments();
            _pending.Enqueue(new Pending(DataTokenType.EndObject));
            if (frame.Node is TomlTable { TableKind: TomlTableKind.Inline } && OwnsLine(frame.Node)) TrailingComment(frame.Node);
            return true;
        }

        if (frame.Items!.MoveNext()) {
            var item = (TomlNode)frame.Items.Current!;
            if (frame.Node is TomlTableArray) {
                LeadingComments(item);
                TrailingComment(item);
            }

            Begin(item);
            return true;
        }

        _frames.Pop();
        _pending.Enqueue(new Pending(DataTokenType.EndArray));
        if (frame.Node is TomlArray && OwnsLine(frame.Node)) TrailingComment(frame.Node);
        return true;
    }

    private void Begin(TomlNode node) {
        _location = node;
        Context.CancellationToken.ThrowIfCancellationRequested();
        switch (node) {
            case TomlTable table:
                _pending.Enqueue(new Pending(DataTokenType.StartObject));
                _frames.Push(new Frame(table, table.GetEnumerator(), null));
                break;
            case TomlTableArray tables:
                _pending.Enqueue(new Pending(DataTokenType.StartArray));
                _frames.Push(new Frame(tables, null, tables.GetEnumerator()));
                break;
            case TomlArray array:
                _pending.Enqueue(new Pending(DataTokenType.StartArray));
                _frames.Push(new Frame(array, null, array.GetEnumerator()));
                break;
            case TomlValue value:
                Scalar(value);
                if (OwnsLine(value)) TrailingComment(value);
                break;
        }
    }

    private void Scalar(TomlValue value) {
        switch (value.Kind) {
            case TomlNodeKind.String:
                _pending.Enqueue(new Pending(DataTokenType.Scalar, DataScalarKind.String, (string)value.Value));
                break;
            case TomlNodeKind.Boolean:
                _pending.Enqueue(new Pending(DataTokenType.Scalar, DataScalarKind.Boolean, (bool)value.Value ? "true" : "false"));
                break;
            case TomlNodeKind.Integer:
                _pending.Enqueue(new Pending(DataTokenType.Scalar, DataScalarKind.Integer, ((long)value.Value).ToString(CultureInfo.InvariantCulture)));
                break;
            case TomlNodeKind.Float:
                _pending.Enqueue(new Pending(DataTokenType.Scalar, DataScalarKind.Float, FloatText(value)));
                break;
            default:
                var date = (TomlDateTime)value.Value;
                DataScalarKind kind = date.Kind switch {
                    TomlDateTimeKind.OffsetDateTime => DataScalarKind.OffsetDateTime,
                    TomlDateTimeKind.LocalDateTime => DataScalarKind.LocalDateTime,
                    TomlDateTimeKind.LocalDate => DataScalarKind.LocalDate,
                    _ => DataScalarKind.LocalTime
                };
                _pending.Enqueue(new Pending(DataTokenType.Scalar, kind, date.ToString()));
                break;
        }
    }

    // The source spelling keeps 1.10 and 5e+22 exact; underscores and a leading '+' are not JSON number syntax.
    internal static string FloatText(TomlValue value) {
        double number = (double)value.Value;
        if (double.IsNaN(number)) return "NaN";
        if (double.IsInfinity(number)) return number > 0 ? "Infinity" : "-Infinity";
        string source = value.ToString();
        var text = new StringBuilder(source.Length);
        foreach (char c in source) {
            if (c == '_' || (c == '+' && text.Length == 0)) continue;
            text.Append(c);
        }

        return text.ToString();
    }

    // Only a key/value pair outside inline tables and arrays ends its own line, so only its trailing comment is its own.
    private static bool OwnsLine(TomlNode node) => node.Parent is TomlTable { TableKind: not TomlTableKind.Inline };

    private void LeadingComments(TomlNode node) {
        IReadOnlyList<string> comments = node.LeadingComments;
        for (int i = 0; i < comments.Count; i++) _pending.Enqueue(new Pending(DataTokenType.Comment, text: CommentText(comments[i])));
        _comments += comments.Count;
    }

    private void TrailingComment(TomlNode node) {
        if (node.TrailingComment is not { } comment) return;
        _pending.Enqueue(new Pending(DataTokenType.Comment, text: CommentText(comment)));
        _comments++;
    }

    // Comment events carry the text without the space after '#', as the YAML and XML sources produce them.
    private static string CommentText(string comment) => comment.Length > 0 && comment[0] == ' ' ? comment.Substring(1) : comment;

    // The engine attaches comments directly above a key or header and after a value. Others, such as a file header
    // separated by a blank line or comments inside arrays, have no node to travel with; count them so the loss is reported.
    private void ReportUnattachedComments() {
        string source = _document.ToString();
        if (source.IndexOf('#') < 0) return;
        var reader = new TomlReader(source, new TomlReaderOptions { IncludeComments = true }, Context.CancellationToken);
        int total = 0;
        while (reader.Read()) {
            if (reader.TokenType == TomlTokenType.Comment) total++;
        }

        if (total > _comments) {
            Context.Report(OfficeDiagnosticCodes.CommentsOmitted, OfficeConversionLossKind.Omission,
                "TOML comments that are not directly above a key or table header, or after a value on its line, were dropped.", null);
        }
    }
}
