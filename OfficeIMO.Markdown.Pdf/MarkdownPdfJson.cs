namespace OfficeIMO.Markdown.Pdf;

internal enum MarkdownPdfJsonValueKind {
    Object,
    Array,
    String,
    Number,
    True,
    False,
    Null
}

internal sealed class MarkdownPdfJsonValue {
    internal const int MaximumInputCharacters = 1_000_000;
    internal const int MaximumNestingDepth = 64;
    internal const int MaximumValueNodes = 100_000;
    private readonly Dictionary<string, MarkdownPdfJsonValue>? _objectValues;
    private readonly List<MarkdownPdfJsonValue>? _arrayValues;

    private MarkdownPdfJsonValue(
        MarkdownPdfJsonValueKind kind,
        Dictionary<string, MarkdownPdfJsonValue>? objectValues = null,
        List<MarkdownPdfJsonValue>? arrayValues = null,
        string? stringValue = null,
        double numberValue = 0D) {
        Kind = kind;
        _objectValues = objectValues;
        _arrayValues = arrayValues;
        StringValue = stringValue;
        NumberValue = numberValue;
    }

    public MarkdownPdfJsonValueKind Kind { get; }

    public string? StringValue { get; }

    public double NumberValue { get; }

    public IReadOnlyList<MarkdownPdfJsonValue> ArrayValues => _arrayValues != null
        ? _arrayValues
        : System.Array.Empty<MarkdownPdfJsonValue>();

    public IReadOnlyDictionary<string, MarkdownPdfJsonValue> ObjectValues => _objectValues != null
        ? _objectValues
        : EmptyObjectValues;

    private static readonly IReadOnlyDictionary<string, MarkdownPdfJsonValue> EmptyObjectValues =
        new Dictionary<string, MarkdownPdfJsonValue>();

    public static MarkdownPdfJsonValue Parse(string json) {
        var parser = new Parser(json);
        return parser.ParseRoot();
    }

    public bool TryGetProperty(string propertyName, out MarkdownPdfJsonValue value) {
        if (_objectValues != null && _objectValues.TryGetValue(propertyName, out MarkdownPdfJsonValue? found)) {
            value = found;
            return true;
        }

        value = Null();
        return false;
    }

    public bool TryGetDouble(out double value) {
        switch (Kind) {
            case MarkdownPdfJsonValueKind.Number:
                value = NumberValue;
                return true;
            case MarkdownPdfJsonValueKind.String:
                return double.TryParse(StringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            default:
                value = 0D;
                return false;
        }
    }

    public string? ReadScalarAsText() {
        switch (Kind) {
            case MarkdownPdfJsonValueKind.String:
                return StringValue;
            case MarkdownPdfJsonValueKind.Number:
                return NumberValue.ToString("0.################", CultureInfo.InvariantCulture);
            case MarkdownPdfJsonValueKind.True:
                return "true";
            case MarkdownPdfJsonValueKind.False:
                return "false";
            case MarkdownPdfJsonValueKind.Array:
                var parts = new List<string>();
                foreach (MarkdownPdfJsonValue item in ArrayValues) {
                    string? text = item.ReadScalarAsText();
                    if (!string.IsNullOrWhiteSpace(text)) {
                        parts.Add(text!);
                    }
                }

                return parts.Count == 0 ? null : string.Join(" ", parts);
            default:
                return null;
        }
    }

    private static MarkdownPdfJsonValue Object(Dictionary<string, MarkdownPdfJsonValue> values) =>
        new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.Object, objectValues: values);

    private static MarkdownPdfJsonValue Array(List<MarkdownPdfJsonValue> values) =>
        new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.Array, arrayValues: values);

    private static MarkdownPdfJsonValue String(string value) =>
        new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.String, stringValue: value);

    private static MarkdownPdfJsonValue Number(double value) =>
        new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.Number, numberValue: value);

    private static MarkdownPdfJsonValue True() => new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.True);

    private static MarkdownPdfJsonValue False() => new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.False);

    private static MarkdownPdfJsonValue Null() => new MarkdownPdfJsonValue(MarkdownPdfJsonValueKind.Null);

    private sealed class Parser {
        private readonly string _text;
        private int _position;
        private int _valueNodes;

        public Parser(string? text) {
            _text = text ?? string.Empty;
            if (_text.Length > MaximumInputCharacters) {
                Throw("Chart JSON exceeds the maximum input size.");
            }
        }

        public MarkdownPdfJsonValue ParseRoot() {
            SkipWhiteSpace();
            MarkdownPdfJsonValue value = ParseValue(0);
            SkipWhiteSpace();
            if (_position != _text.Length) {
                Throw("Unexpected trailing JSON content.");
            }

            return value;
        }

        private MarkdownPdfJsonValue ParseValue(int depth) {
            SkipWhiteSpace();
            if (depth > MaximumNestingDepth) Throw("Chart JSON exceeds the maximum nesting depth.");
            if (++_valueNodes > MaximumValueNodes) Throw("Chart JSON exceeds the maximum value count.");
            if (_position >= _text.Length) {
                Throw("Unexpected end of JSON.");
            }

            char ch = _text[_position];
            switch (ch) {
                case '{':
                    return ParseObject(depth);
                case '[':
                    return ParseArray(depth);
                case '"':
                    return String(ParseString());
                case 't':
                    ConsumeLiteral("true");
                    return True();
                case 'f':
                    ConsumeLiteral("false");
                    return False();
                case 'n':
                    ConsumeLiteral("null");
                    return Null();
                default:
                    if (ch == '-' || char.IsDigit(ch)) {
                        return Number(ParseNumber());
                    }

                    Throw("Unexpected JSON token.");
                    return Null();
            }
        }

        private MarkdownPdfJsonValue ParseObject(int depth) {
            Expect('{');
            var values = new Dictionary<string, MarkdownPdfJsonValue>(StringComparer.OrdinalIgnoreCase);
            SkipWhiteSpace();
            if (TryConsume('}')) {
                return Object(values);
            }

            while (true) {
                SkipWhiteSpace();
                if (!Peek('"')) {
                    Throw("JSON object property names must be strings.");
                }

                string propertyName = ParseString();
                SkipWhiteSpace();
                Expect(':');
                values[propertyName] = ParseValue(depth + 1);
                SkipWhiteSpace();
                if (TryConsume('}')) {
                    return Object(values);
                }

                Expect(',');
            }
        }

        private MarkdownPdfJsonValue ParseArray(int depth) {
            Expect('[');
            var values = new List<MarkdownPdfJsonValue>();
            SkipWhiteSpace();
            if (TryConsume(']')) {
                return Array(values);
            }

            while (true) {
                values.Add(ParseValue(depth + 1));
                SkipWhiteSpace();
                if (TryConsume(']')) {
                    return Array(values);
                }

                Expect(',');
            }
        }

        private string ParseString() {
            Expect('"');
            var builder = new StringBuilder();
            while (_position < _text.Length) {
                char ch = _text[_position++];
                if (ch == '"') {
                    return builder.ToString();
                }

                if (ch != '\\') {
                    builder.Append(ch);
                    continue;
                }

                if (_position >= _text.Length) {
                    Throw("Unterminated JSON string escape.");
                }

                char escaped = _text[_position++];
                switch (escaped) {
                    case '"':
                    case '\\':
                    case '/':
                        builder.Append(escaped);
                        break;
                    case 'b':
                        builder.Append('\b');
                        break;
                    case 'f':
                        builder.Append('\f');
                        break;
                    case 'n':
                        builder.Append('\n');
                        break;
                    case 'r':
                        builder.Append('\r');
                        break;
                    case 't':
                        builder.Append('\t');
                        break;
                    case 'u':
                        builder.Append(ParseUnicodeEscape());
                        break;
                    default:
                        Throw("Unsupported JSON string escape.");
                        break;
                }
            }

            Throw("Unterminated JSON string.");
            return string.Empty;
        }

        private char ParseUnicodeEscape() {
            if (_position + 4 > _text.Length) {
                Throw("Incomplete JSON unicode escape.");
            }

            int value = 0;
            for (int i = 0; i < 4; i++) {
                char ch = _text[_position++];
                int digit = HexValue(ch);
                if (digit < 0) {
                    Throw("Invalid JSON unicode escape.");
                }

                value = value * 16 + digit;
            }

            return (char)value;
        }

        private double ParseNumber() {
            int start = _position;
            if (Peek('-')) {
                _position++;
            }

            if (Peek('0')) {
                _position++;
            } else {
                RequireDigit();
                while (_position < _text.Length && char.IsDigit(_text[_position])) {
                    _position++;
                }
            }

            if (Peek('.')) {
                _position++;
                RequireDigit();
                while (_position < _text.Length && char.IsDigit(_text[_position])) {
                    _position++;
                }
            }

            if (_position < _text.Length && (_text[_position] == 'e' || _text[_position] == 'E')) {
                _position++;
                if (_position < _text.Length && (_text[_position] == '+' || _text[_position] == '-')) {
                    _position++;
                }

                RequireDigit();
                while (_position < _text.Length && char.IsDigit(_text[_position])) {
                    _position++;
                }
            }

            string token = _text.Substring(start, _position - start);
            if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                double.IsNaN(value) ||
                double.IsInfinity(value)) {
                Throw("Invalid JSON number.");
            }

            return value;
        }

        private void ConsumeLiteral(string literal) {
            if (_position + literal.Length > _text.Length ||
                string.Compare(_text, _position, literal, 0, literal.Length, StringComparison.Ordinal) != 0) {
                Throw("Invalid JSON literal.");
            }

            _position += literal.Length;
        }

        private void RequireDigit() {
            if (_position >= _text.Length || !char.IsDigit(_text[_position])) {
                Throw("JSON number expected a digit.");
            }
        }

        private void Expect(char expected) {
            SkipWhiteSpace();
            if (!TryConsume(expected)) {
                Throw("Expected '" + expected + "'.");
            }
        }

        private bool TryConsume(char expected) {
            if (Peek(expected)) {
                _position++;
                return true;
            }

            return false;
        }

        private bool Peek(char expected) => _position < _text.Length && _text[_position] == expected;

        private void SkipWhiteSpace() {
            while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) {
                _position++;
            }
        }

        private static int HexValue(char ch) {
            if (ch >= '0' && ch <= '9') {
                return ch - '0';
            }

            if (ch >= 'a' && ch <= 'f') {
                return 10 + ch - 'a';
            }

            if (ch >= 'A' && ch <= 'F') {
                return 10 + ch - 'A';
            }

            return -1;
        }

        private static void Throw(string message) => throw new FormatException(message);
    }
}
