using System.Globalization;
using System.Text;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>
/// Small JSON reader for the suite's <c>in.json</c> files and the asset manifests.
/// It accepts a stream of top-level values, one per YAML document, separated by whitespace.
/// </summary>
public static class YamlJsonReader {
    /// <summary>Reads every top-level value in the text.</summary>
    public static IReadOnlyList<YamlConformanceNode> ReadStream(string json) {
        if (json == null) throw new ArgumentNullException(nameof(json));
        var scanner = new Scanner(json);
        var values = new List<YamlConformanceNode>();
        scanner.SkipWhitespace();
        while (!scanner.AtEnd) {
            values.Add(scanner.ReadValue(0));
            scanner.SkipWhitespace();
        }

        return values;
    }

    /// <summary>Reads exactly one top-level value.</summary>
    public static YamlConformanceNode Read(string json) {
        IReadOnlyList<YamlConformanceNode> values = ReadStream(json);
        if (values.Count != 1) throw new FormatException("Expected one JSON value but found " + values.Count + ".");
        return values[0];
    }

    private sealed class Scanner {
        private const int MaxDepth = 256;
        private readonly string _text;
        private int _position;

        public Scanner(string text) {
            _text = text;
            // A leading byte order mark is not JSON content.
            if (_text.Length > 0 && _text[0] == '﻿') _position = 1;
        }

        public bool AtEnd => _position >= _text.Length;

        public void SkipWhitespace() {
            while (_position < _text.Length) {
                char c = _text[_position];
                if (c != ' ' && c != '\t' && c != '\n' && c != '\r') break;
                _position++;
            }
        }

        public YamlConformanceNode ReadValue(int depth) {
            if (depth > MaxDepth) throw Error("JSON nesting is too deep");
            SkipWhitespace();
            if (AtEnd) throw Error("Unexpected end of JSON");
            char c = _text[_position];
            switch (c) {
                case '{': return ReadObject(depth);
                case '[': return ReadArray(depth);
                case '"': return YamlConformanceNode.String(ReadString());
                case 't': ExpectWord("true"); return YamlConformanceNode.Boolean(true);
                case 'f': ExpectWord("false"); return YamlConformanceNode.Boolean(false);
                case 'n': ExpectWord("null"); return YamlConformanceNode.Null;
                default:
                    if (c == '-' || (c >= '0' && c <= '9')) return ReadNumber();
                    throw Error("Unexpected character '" + c + "'");
            }
        }

        private YamlConformanceNode ReadObject(int depth) {
            _position++;
            var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>();
            SkipWhitespace();
            if (Peek() == '}') {
                _position++;
                return YamlConformanceNode.Mapping(entries);
            }

            while (true) {
                SkipWhitespace();
                if (Peek() != '"') throw Error("Expected a property name");
                YamlConformanceNode key = YamlConformanceNode.String(ReadString());
                SkipWhitespace();
                Expect(':');
                YamlConformanceNode value = ReadValue(depth + 1);
                entries.Add(new KeyValuePair<YamlConformanceNode, YamlConformanceNode>(key, value));
                SkipWhitespace();
                char next = Peek();
                _position++;
                if (next == '}') return YamlConformanceNode.Mapping(entries);
                if (next != ',') throw Error("Expected ',' or '}'", _position - 1);
            }
        }

        private YamlConformanceNode ReadArray(int depth) {
            _position++;
            var items = new List<YamlConformanceNode>();
            SkipWhitespace();
            if (Peek() == ']') {
                _position++;
                return YamlConformanceNode.Sequence(items);
            }

            while (true) {
                items.Add(ReadValue(depth + 1));
                SkipWhitespace();
                char next = Peek();
                _position++;
                if (next == ']') return YamlConformanceNode.Sequence(items);
                if (next != ',') throw Error("Expected ',' or ']'", _position - 1);
            }
        }

        private string ReadString() {
            Expect('"');
            var value = new StringBuilder();
            while (true) {
                if (AtEnd) throw Error("Unterminated string");
                char c = _text[_position++];
                if (c == '"') return value.ToString();
                if (c < ' ') throw Error("Control character in string", _position - 1);
                if (c != '\\') {
                    value.Append(c);
                    continue;
                }

                if (AtEnd) throw Error("Unterminated escape");
                char escape = _text[_position++];
                switch (escape) {
                    case '"': value.Append('"'); break;
                    case '\\': value.Append('\\'); break;
                    case '/': value.Append('/'); break;
                    case 'b': value.Append('\b'); break;
                    case 'f': value.Append('\f'); break;
                    case 'n': value.Append('\n'); break;
                    case 'r': value.Append('\r'); break;
                    case 't': value.Append('\t'); break;
                    // Surrogate pairs arrive as two escapes; appending both code units rebuilds the pair.
                    case 'u': value.Append((char)ReadHex4()); break;
                    default: throw Error("Invalid escape '\\" + escape + "'", _position - 1);
                }
            }
        }

        private int ReadHex4() {
            if (_position + 4 > _text.Length) throw Error("Truncated \\u escape");
            int result = 0;
            for (int i = 0; i < 4; i++) {
                char c = _text[_position++];
                int digit = c >= '0' && c <= '9' ? c - '0'
                    : c >= 'a' && c <= 'f' ? c - 'a' + 10
                    : c >= 'A' && c <= 'F' ? c - 'A' + 10
                    : -1;
                if (digit < 0) throw Error("Invalid hex digit '" + c + "'", _position - 1);
                result = (result << 4) | digit;
            }

            return result;
        }

        // JSON number grammar: -?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?
        private YamlConformanceNode ReadNumber() {
            int start = _position;
            bool isFloat = false;
            if (Peek() == '-') _position++;
            if (Peek() == '0') {
                _position++;
            } else {
                if (!IsDigit(Peek())) throw Error("Invalid number");
                SkipDigits();
            }

            if (Peek() == '.') {
                isFloat = true;
                _position++;
                if (!IsDigit(Peek())) throw Error("Invalid fraction");
                SkipDigits();
            }

            if (Peek() == 'e' || Peek() == 'E') {
                isFloat = true;
                _position++;
                if (Peek() == '+' || Peek() == '-') _position++;
                if (!IsDigit(Peek())) throw Error("Invalid exponent");
                SkipDigits();
            }

            string text = _text.Substring(start, _position - start);
            return isFloat ? YamlConformanceNode.Float(text) : YamlConformanceNode.Integer(text);
        }

        private void SkipDigits() {
            while (IsDigit(Peek())) _position++;
        }

        private static bool IsDigit(char c) => c >= '0' && c <= '9';

        private char Peek() => _position < _text.Length ? _text[_position] : '\0';

        private void Expect(char expected) {
            if (Peek() != expected) throw Error("Expected '" + expected + "'");
            _position++;
        }

        private void ExpectWord(string word) {
            if (string.CompareOrdinal(_text, _position, word, 0, word.Length) != 0) throw Error("Expected '" + word + "'");
            _position += word.Length;
        }

        private FormatException Error(string message) => Error(message, _position);

        private static FormatException Error(string message, int position) =>
            new(message + " at offset " + position.ToString(CultureInfo.InvariantCulture) + ".");
    }
}
