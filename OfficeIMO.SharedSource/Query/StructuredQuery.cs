using System.Globalization;
using System.Text;

namespace OfficeIMO.Query;

/// <summary>What a node is to a query: a scalar, a mapping (YAML mapping, TOML table) or a sequence (YAML sequence, TOML array or array of tables).</summary>
internal enum QueryNodeKind : byte {
    Scalar,
    Mapping,
    Sequence
}

/// <summary>The type a scalar compares as in a filter.</summary>
internal enum QueryScalarKind : byte {
    Null,
    Boolean,
    Number,
    String
}

/// <summary>A scalar value as filters compare it.</summary>
internal readonly struct QueryScalar {
    private QueryScalar(QueryScalarKind kind, bool boolean, double number, string? text) {
        Kind = kind;
        Boolean = boolean;
        Number = number;
        Text = text;
    }

    public QueryScalarKind Kind { get; }
    public bool Boolean { get; }
    public double Number { get; }
    public string? Text { get; }

    public static QueryScalar Null => new QueryScalar(QueryScalarKind.Null, false, 0, null);
    public static QueryScalar FromBoolean(bool value) => new QueryScalar(QueryScalarKind.Boolean, value, 0, null);
    public static QueryScalar FromNumber(double value) => new QueryScalar(QueryScalarKind.Number, false, value, null);
    public static QueryScalar FromString(string value) => new QueryScalar(QueryScalarKind.String, false, 0, value);
}

/// <summary>A child of a node: a mapping value with its key text, or a sequence item with its index.</summary>
internal readonly struct QueryChild<TNode> where TNode : class {
    public QueryChild(string key, TNode node) {
        Key = key;
        Index = -1;
        Node = node;
    }

    public QueryChild(int index, TNode node) {
        Key = null;
        Index = index;
        Node = node;
    }

    public string? Key { get; }
    public int Index { get; }
    public TNode Node { get; }
}

/// <summary>A node a query selected, with the path that addresses it in the format's <c>Select</c> grammar.</summary>
internal readonly struct QueryMatch<TNode> where TNode : class {
    public QueryMatch(TNode node, string path) {
        Node = node;
        Path = path;
    }

    public TNode Node { get; }
    public string Path { get; }
}

/// <summary>Gives the query evaluator a format-neutral view of one document model.</summary>
internal interface IQueryAdapter<TNode> where TNode : class {
    /// <summary>Whether the node stands for another written elsewhere (a YAML alias); recursive descent passes over it.</summary>
    bool IsReference(TNode node);

    /// <summary>The node's kind, looking through YAML aliases.</summary>
    QueryNodeKind GetKind(TNode node);

    /// <summary>The value a key step selects, with the format's <c>Select</c> rules (YAML merge keys included), or null.</summary>
    TNode? GetMember(TNode node, string key);

    /// <summary>The number of items in a sequence.</summary>
    int GetCount(TNode node);

    /// <summary>The item of a sequence.</summary>
    TNode GetItem(TNode node, int index);

    /// <summary>
    /// Adds a mapping's values or a sequence's items. With <paramref name="written"/>, mapping values are the entries as
    /// written (YAML merge keys left out) and aliases have no children, so recursive descent visits each node once,
    /// where it is written; otherwise mapping values are the ones key steps reach, YAML merged keys included.
    /// </summary>
    void AddChildren(TNode node, bool written, List<QueryChild<TNode>> children);

    /// <summary>The scalar value a filter compares, or false for collections.</summary>
    bool TryGetScalar(TNode node, out QueryScalar value);

    /// <summary>Appends a key step to a path in the format's <c>Select</c> grammar.</summary>
    string AppendKey(string path, string key);

    /// <summary>Appends an index step to a path.</summary>
    string AppendIndex(string path, int index);
}

/// <summary>The kind of a query step.</summary>
internal enum QueryStepKind : byte {
    Key,
    Index,
    Wildcard,
    Filter
}

/// <summary>One step of a query, optionally applied to every descendant (<c>..</c>).</summary>
internal sealed class QueryStep {
    public QueryStepKind Kind { get; set; }
    public bool Descendants { get; set; }
    public string? Key { get; set; }
    public long Index { get; set; }
    public QueryExpression? Filter { get; set; }
}

/// <summary>
/// A parsed query over YAML, TOML or JSON nodes: the <c>Select</c> path grammar plus wildcards (<c>*</c>, <c>[*]</c>),
/// recursive descent (<c>..name</c>, <c>..*</c>, <c>..[0]</c>), negative indexes (<c>[-1]</c>) and filters
/// (<c>[?(@.name == 'web')]</c>, <c>[?@.port &gt;= 8000 &amp;&amp; @.enabled]</c>).
/// </summary>
internal sealed class StructuredQuery {
    private readonly string _text;
    private readonly string _format;
    private int _i;

    private StructuredQuery(string text, string format) {
        _text = text;
        _format = format;
    }

    /// <summary>Whether the query starts with <c>$</c>.</summary>
    public bool Rooted { get; private set; }

    public List<QueryStep> Steps { get; } = new List<QueryStep>();

    /// <summary>Parses a query; <paramref name="format"/> names the format in error messages ("YAML", "TOML").</summary>
    public static StructuredQuery Parse(string query, string format) {
        if (query is null) throw new ArgumentNullException(nameof(query));
        var parsed = new StructuredQuery(query, format);
        parsed.ParseSteps();
        return parsed;
    }

    /// <summary>Evaluates the steps from <paramref name="first"/> on, starting from <paramref name="start"/>.</summary>
    public List<QueryMatch<TNode>> Evaluate<TNode>(IQueryAdapter<TNode> adapter, List<QueryMatch<TNode>> start, int first) where TNode : class {
        List<QueryMatch<TNode>> current = start;
        var children = new List<QueryChild<TNode>>();
        for (int s = first; s < Steps.Count; s++) {
            QueryStep step = Steps[s];
            var next = new List<QueryMatch<TNode>>();
            foreach (QueryMatch<TNode> match in current) {
                if (step.Descendants) {
                    // Pre-order over the nodes as written, so results come in document order.
                    // An alias is reached where its anchor is written, so steps skip it below the starting node.
                    var stack = new Stack<QueryMatch<TNode>>();
                    stack.Push(match);
                    while (stack.Count > 0) {
                        QueryMatch<TNode> node = stack.Pop();
                        if (ReferenceEquals(node.Node, match.Node) || !adapter.IsReference(node.Node)) Apply(adapter, step, node, next, children);
                        children.Clear();
                        if (adapter.GetKind(node.Node) == QueryNodeKind.Scalar) continue;
                        adapter.AddChildren(node.Node, written: true, children);
                        for (int c = children.Count - 1; c >= 0; c--) stack.Push(Child(adapter, node.Path, children[c]));
                        children.Clear();
                    }
                } else {
                    Apply(adapter, step, match, next, children);
                }
            }
            current = next;
        }
        return current;
    }

    private static QueryMatch<TNode> Child<TNode>(IQueryAdapter<TNode> adapter, string path, QueryChild<TNode> child) where TNode : class =>
        new QueryMatch<TNode>(child.Node, child.Key is null ? adapter.AppendIndex(path, child.Index) : adapter.AppendKey(path, child.Key));

    private static void Apply<TNode>(IQueryAdapter<TNode> adapter, QueryStep step, QueryMatch<TNode> match, List<QueryMatch<TNode>> results, List<QueryChild<TNode>> children) where TNode : class {
        QueryNodeKind kind = adapter.GetKind(match.Node);
        switch (step.Kind) {
            case QueryStepKind.Key:
                if (kind == QueryNodeKind.Mapping && adapter.GetMember(match.Node, step.Key!) is { } value) {
                    results.Add(new QueryMatch<TNode>(value, adapter.AppendKey(match.Path, step.Key!)));
                }
                break;
            case QueryStepKind.Index:
                if (kind == QueryNodeKind.Sequence) {
                    int count = adapter.GetCount(match.Node);
                    long index = step.Index < 0 ? count + step.Index : step.Index;
                    if (index >= 0 && index < count) results.Add(new QueryMatch<TNode>(adapter.GetItem(match.Node, (int)index), adapter.AppendIndex(match.Path, (int)index)));
                }
                break;
            default:
                if (kind == QueryNodeKind.Scalar) break;
                children.Clear();
                adapter.AddChildren(match.Node, written: false, children);
                foreach (QueryChild<TNode> child in children) {
                    if (step.Kind == QueryStepKind.Filter && !step.Filter!.IsTrue(adapter, child.Node)) continue;
                    results.Add(Child(adapter, match.Path, child));
                }
                children.Clear();
                break;
        }
    }

    private void ParseSteps() {
        string text = _text;
        if (_i < text.Length && text[_i] == '$') {
            Rooted = true;
            _i++;
        } else if (_i < text.Length && text[_i] != '.' && text[_i] != '[') {
            Steps.Add(NameStep(ReadName(), descendants: false));
        }
        while (_i < text.Length) {
            char ch = text[_i];
            if (ch == '.') {
                _i++;
                bool descendants = _i < text.Length && text[_i] == '.';
                if (descendants) {
                    _i++;
                    if (_i < text.Length && text[_i] == '[') {
                        _i++;
                        Steps.Add(ReadBracket(descendants: true));
                        continue;
                    }
                }
                Steps.Add(NameStep(ReadName(), descendants));
            } else if (ch == '[') {
                _i++;
                Steps.Add(ReadBracket(descendants: false));
            } else {
                throw Malformed(_i, "expected '.' or '['");
            }
        }
    }

    private static QueryStep NameStep(string name, bool descendants) =>
        name == "*" ? new QueryStep { Kind = QueryStepKind.Wildcard, Descendants = descendants } : new QueryStep { Kind = QueryStepKind.Key, Key = name, Descendants = descendants };

    /// <summary>A name runs to the next <c>.</c> or <c>[</c>, as in <c>Select</c> paths.</summary>
    private string ReadName() {
        int start = _i;
        while (_i < _text.Length && _text[_i] != '.' && _text[_i] != '[') _i++;
        if (_i == start) throw Malformed(start, "expected a key name");
        return _text.Substring(start, _i - start);
    }

    /// <summary>Reads what follows <c>[</c>: <c>*]</c>, a quoted key, an index or a filter.</summary>
    private QueryStep ReadBracket(bool descendants) {
        string text = _text;
        if (_i >= text.Length) throw Malformed(_i, "expected an index, a quoted key, '*' or a filter");
        char ch = text[_i];
        QueryStep step;
        if (ch == '*') {
            _i++;
            step = new QueryStep { Kind = QueryStepKind.Wildcard, Descendants = descendants };
        } else if (ch == '"' || ch == '\'') {
            step = new QueryStep { Kind = QueryStepKind.Key, Key = ReadQuoted(), Descendants = descendants };
        } else if (ch == '?') {
            _i++;
            QueryExpression filter = ReadOr();
            SkipSpace();
            step = new QueryStep { Kind = QueryStepKind.Filter, Filter = filter, Descendants = descendants };
        } else {
            step = new QueryStep { Kind = QueryStepKind.Index, Index = ReadIndex(), Descendants = descendants };
        }
        Expect(']');
        return step;
    }

    /// <summary>Reads an optionally negative index. Indexes beyond the int range simply match nothing.</summary>
    private long ReadIndex() {
        int start = _i;
        bool negative = _i < _text.Length && _text[_i] == '-';
        if (negative) _i++;
        int digits = _i;
        long value = 0;
        while (_i < _text.Length && _text[_i] >= '0' && _text[_i] <= '9') {
            if (value < int.MaxValue) value = value * 10 + (_text[_i] - '0');
            _i++;
        }
        if (_i == digits) throw Malformed(start, "expected an index, a quoted key, '*' or a filter");
        return negative ? -value : value;
    }

    /// <summary>Reads a single- or double-quoted string with JSON escapes (and <c>\'</c>).</summary>
    private string ReadQuoted() {
        string text = _text;
        int open = _i;
        char quote = text[_i++];
        var value = new StringBuilder();
        while (true) {
            if (_i >= text.Length) throw Malformed(open, "the quoted text is not closed");
            char ch = text[_i++];
            if (ch == quote) return value.ToString();
            if (ch != '\\') {
                value.Append(ch);
                continue;
            }
            if (_i >= text.Length) throw Malformed(_i, "the escape is incomplete");
            char escape = text[_i++];
            switch (escape) {
                case '"': case '\'': case '\\': case '/': value.Append(escape); break;
                case 'b': value.Append('\b'); break;
                case 'f': value.Append('\f'); break;
                case 'n': value.Append('\n'); break;
                case 'r': value.Append('\r'); break;
                case 't': value.Append('\t'); break;
                case 'u':
                    int code = 0;
                    for (int k = 0; k < 4; k++) {
                        int digit = _i < text.Length ? HexValue(text[_i]) : -1;
                        if (digit < 0) throw Malformed(_i, "expected four hexadecimal digits after \\u");
                        code = code * 16 + digit;
                        _i++;
                    }
                    value.Append((char)code);
                    break;
                default:
                    throw Malformed(_i - 1, "unknown escape '\\" + escape + "'");
            }
        }
    }

    private static int HexValue(char c) =>
        c >= '0' && c <= '9' ? c - '0' : c >= 'a' && c <= 'f' ? c - 'a' + 10 : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;

    private QueryExpression ReadOr() {
        QueryExpression left = ReadAnd();
        while (TryRead("||")) left = new QueryExpression(QueryOperator.Or, left, ReadAnd());
        return left;
    }

    private QueryExpression ReadAnd() {
        QueryExpression left = ReadUnary();
        while (TryRead("&&")) left = new QueryExpression(QueryOperator.And, left, ReadUnary());
        return left;
    }

    private QueryExpression ReadUnary() {
        SkipSpace();
        if (_i < _text.Length && _text[_i] == '!' && (_i + 1 >= _text.Length || _text[_i + 1] != '=')) {
            _i++;
            return new QueryExpression(QueryOperator.Not, ReadUnary(), null);
        }
        if (_i < _text.Length && _text[_i] == '(') {
            _i++;
            QueryExpression inner = ReadOr();
            SkipSpace();
            Expect(')');
            return inner;
        }
        int start = _i;
        QueryOperand left = ReadOperand();
        SkipSpace();
        QueryOperator op = ReadComparison();
        if (op == QueryOperator.Exists) {
            if (left.Path is null) throw Malformed(start, "a literal alone is not a condition; compare it or test a '@' path");
            return new QueryExpression(left);
        }
        SkipSpace();
        return new QueryExpression(op, left, ReadOperand());
    }

    private QueryOperator ReadComparison() {
        if (TryRead("==")) return QueryOperator.Equal;
        if (TryRead("!=")) return QueryOperator.NotEqual;
        if (TryRead("<=")) return QueryOperator.LessOrEqual;
        if (TryRead(">=")) return QueryOperator.GreaterOrEqual;
        if (TryRead("<")) return QueryOperator.Less;
        if (TryRead(">")) return QueryOperator.Greater;
        return QueryOperator.Exists;
    }

    private QueryOperand ReadOperand() {
        SkipSpace();
        string text = _text;
        if (_i >= text.Length) throw Malformed(_i, "expected a value or a '@' path");
        char ch = text[_i];
        if (ch == '@') {
            _i++;
            return QueryOperand.FromPath(ReadRelativePath());
        }
        if (ch == '"' || ch == '\'') return QueryOperand.FromScalar(QueryScalar.FromString(ReadQuoted()));
        if (ch == '-' || (ch >= '0' && ch <= '9')) {
            int start = _i;
            _i++;
            while (_i < text.Length && (char.IsDigit(text[_i]) || text[_i] == '.' || text[_i] == 'e' || text[_i] == 'E' || ((text[_i] == '+' || text[_i] == '-') && (text[_i - 1] == 'e' || text[_i - 1] == 'E')))) _i++;
            if (!double.TryParse(text.Substring(start, _i - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) throw Malformed(start, "the number is not valid");
            return QueryOperand.FromScalar(QueryScalar.FromNumber(number));
        }
        if (TryWord("true")) return QueryOperand.FromScalar(QueryScalar.FromBoolean(true));
        if (TryWord("false")) return QueryOperand.FromScalar(QueryScalar.FromBoolean(false));
        if (TryWord("null")) return QueryOperand.FromScalar(QueryScalar.Null);
        if (ch == '$') throw Malformed(_i, "filters compare '@' paths and literals; '$' paths are not supported in filters");
        throw Malformed(_i, "expected a value or a '@' path");
    }

    /// <summary>Reads the key and index steps after <c>@</c>; names stop at spaces and operator characters too.</summary>
    private List<QueryStep> ReadRelativePath() {
        string text = _text;
        var steps = new List<QueryStep>();
        while (_i < text.Length) {
            char ch = text[_i];
            if (ch == '.') {
                _i++;
                int start = _i;
                while (_i < text.Length && !IsFilterDelimiter(text[_i])) _i++;
                if (_i == start) throw Malformed(start, "expected a key name");
                steps.Add(new QueryStep { Kind = QueryStepKind.Key, Key = text.Substring(start, _i - start) });
            } else if (ch == '[') {
                _i++;
                if (_i < text.Length && (text[_i] == '"' || text[_i] == '\'')) {
                    steps.Add(new QueryStep { Kind = QueryStepKind.Key, Key = ReadQuoted() });
                } else {
                    steps.Add(new QueryStep { Kind = QueryStepKind.Index, Index = ReadIndex() });
                }
                Expect(']');
            } else {
                break;
            }
        }
        return steps;
    }

    private static bool IsFilterDelimiter(char c) =>
        c == '.' || c == '[' || c == ']' || c == '(' || c == ')' || c == '=' || c == '!' || c == '<' || c == '>' || c == '&' || c == '|' || char.IsWhiteSpace(c);

    private bool TryWord(string word) {
        if (string.CompareOrdinal(_text, _i, word, 0, word.Length) != 0) return false;
        int end = _i + word.Length;
        if (end < _text.Length && char.IsLetterOrDigit(_text[end])) return false;
        _i = end;
        return true;
    }

    private bool TryRead(string token) {
        SkipSpace();
        if (string.CompareOrdinal(_text, _i, token, 0, token.Length) != 0) return false;
        _i += token.Length;
        return true;
    }

    private void SkipSpace() {
        while (_i < _text.Length && (_text[_i] == ' ' || _text[_i] == '\t')) _i++;
    }

    private void Expect(char ch) {
        if (_i >= _text.Length || _text[_i] != ch) throw Malformed(_i, "expected '" + ch + "'");
        _i++;
    }

    private ArgumentException Malformed(int position, string reason) =>
        new ArgumentException("The " + _format + " query '" + _text + "' is malformed at position " + position.ToString(CultureInfo.InvariantCulture) + ": " + reason + ".", "query");
}

internal enum QueryOperator : byte {
    Exists,
    Equal,
    NotEqual,
    Less,
    LessOrEqual,
    Greater,
    GreaterOrEqual,
    And,
    Or,
    Not
}

/// <summary>A filter operand: a <c>@</c> path or a literal.</summary>
internal sealed class QueryOperand {
    public List<QueryStep>? Path { get; private set; }
    public QueryScalar Literal { get; private set; }

    public static QueryOperand FromPath(List<QueryStep> path) => new QueryOperand { Path = path };
    public static QueryOperand FromScalar(QueryScalar value) => new QueryOperand { Literal = value };
}

/// <summary>A filter condition.</summary>
internal sealed class QueryExpression {
    private readonly QueryOperator _operator;
    private readonly QueryExpression? _left;
    private readonly QueryExpression? _right;
    private readonly QueryOperand? _leftOperand;
    private readonly QueryOperand? _rightOperand;

    public QueryExpression(QueryOperand path) {
        _operator = QueryOperator.Exists;
        _leftOperand = path;
    }

    public QueryExpression(QueryOperator op, QueryExpression left, QueryExpression? right) {
        _operator = op;
        _left = left;
        _right = right;
    }

    public QueryExpression(QueryOperator op, QueryOperand left, QueryOperand right) {
        _operator = op;
        _leftOperand = left;
        _rightOperand = right;
    }

    public bool IsTrue<TNode>(IQueryAdapter<TNode> adapter, TNode current) where TNode : class {
        switch (_operator) {
            case QueryOperator.And: return _left!.IsTrue(adapter, current) && _right!.IsTrue(adapter, current);
            case QueryOperator.Or: return _left!.IsTrue(adapter, current) || _right!.IsTrue(adapter, current);
            case QueryOperator.Not: return !_left!.IsTrue(adapter, current);
            case QueryOperator.Exists: return Resolve(adapter, current, _leftOperand!.Path!) is not null;
        }
        Value left = Evaluate(adapter, current, _leftOperand!);
        Value right = Evaluate(adapter, current, _rightOperand!);
        switch (_operator) {
            case QueryOperator.Equal: return AreEqual(left, right);
            case QueryOperator.NotEqual: return !AreEqual(left, right);
            case QueryOperator.Less: return Compare(left, right) is < 0;
            case QueryOperator.LessOrEqual: return Compare(left, right) is <= 0 || AreEqual(left, right);
            case QueryOperator.Greater: return Compare(left, right) is > 0;
            default: return Compare(left, right) is >= 0 || AreEqual(left, right);
        }
    }

    /// <summary>What an operand evaluates to: nothing, a scalar, or a collection (which only equals nothing else).</summary>
    private readonly struct Value {
        public Value(bool found, bool scalar, QueryScalar value) {
            Found = found;
            IsScalar = scalar;
            Scalar = value;
        }

        public bool Found { get; }
        public bool IsScalar { get; }
        public QueryScalar Scalar { get; }
    }

    private static Value Evaluate<TNode>(IQueryAdapter<TNode> adapter, TNode current, QueryOperand operand) where TNode : class {
        if (operand.Path is null) return new Value(true, true, operand.Literal);
        TNode? node = Resolve(adapter, current, operand.Path);
        if (node is null) return default;
        return adapter.TryGetScalar(node, out QueryScalar scalar) ? new Value(true, true, scalar) : new Value(true, false, default);
    }

    private static TNode? Resolve<TNode>(IQueryAdapter<TNode> adapter, TNode current, List<QueryStep> path) where TNode : class {
        TNode? node = current;
        foreach (QueryStep step in path) {
            if (node is null) return null;
            QueryNodeKind kind = adapter.GetKind(node);
            if (step.Kind == QueryStepKind.Key) {
                node = kind == QueryNodeKind.Mapping ? adapter.GetMember(node, step.Key!) : null;
            } else if (kind == QueryNodeKind.Sequence) {
                int count = adapter.GetCount(node);
                long index = step.Index < 0 ? count + step.Index : step.Index;
                node = index >= 0 && index < count ? adapter.GetItem(node, (int)index) : null;
            } else {
                node = null;
            }
        }
        return node;
    }

    /// <summary>Missing equals missing; scalars of different kinds and collections are never equal.</summary>
    private static bool AreEqual(Value left, Value right) {
        if (!left.Found || !right.Found) return !left.Found && !right.Found;
        if (!left.IsScalar || !right.IsScalar) return false;
        QueryScalar a = left.Scalar, b = right.Scalar;
        if (a.Kind != b.Kind) return false;
        return a.Kind switch {
            QueryScalarKind.Null => true,
            QueryScalarKind.Boolean => a.Boolean == b.Boolean,
            QueryScalarKind.Number => a.Number.Equals(b.Number),
            _ => string.Equals(a.Text, b.Text, StringComparison.Ordinal)
        };
    }

    /// <summary>Orders two numbers or two strings (ordinally); null for anything else.</summary>
    private static int? Compare(Value left, Value right) {
        if (!left.Found || !right.Found || !left.IsScalar || !right.IsScalar) return null;
        QueryScalar a = left.Scalar, b = right.Scalar;
        if (a.Kind == QueryScalarKind.Number && b.Kind == QueryScalarKind.Number) {
            if (double.IsNaN(a.Number) || double.IsNaN(b.Number)) return null;
            return a.Number.CompareTo(b.Number);
        }
        if (a.Kind == QueryScalarKind.String && b.Kind == QueryScalarKind.String) return string.CompareOrdinal(a.Text, b.Text);
        return null;
    }
}
