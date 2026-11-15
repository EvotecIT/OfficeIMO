using System.Globalization;
using System.Text;
namespace OfficeIMO.Json;

/// <summary>A strict RFC 8259 parser that records the line and column of every value.</summary>
internal sealed class JsonDataParser {
    internal const int MaxDepth = 512;

    private readonly string _text;
    private int _pos;
    private int _line = 1;
    private int _lineStart;

    public JsonDataParser(string text) {
        _text = text;
        // A leading byte-order mark is not part of the value.
        if (_text.Length > 0 && _text[0] == '﻿') _pos = 1;
    }

    public JsonData ParseDocument() {
        SkipWhitespace();
        JsonData value = ParseValue(0);
        SkipWhitespace();
        if (_pos < _text.Length) throw Error("Unexpected text after the JSON value");
        return value;
    }

    private JsonData ParseValue(int depth) {
        if (depth > MaxDepth) throw Error("JSON nests deeper than " + MaxDepth.ToString(CultureInfo.InvariantCulture) + " levels");
        if (_pos >= _text.Length) throw Error("Unexpected end of JSON");
        var at = new JsonDataLocation(_line, _pos - _lineStart + 1);
        char c = _text[_pos];
        switch (c) {
            case '{': return ParseObject(depth, at);
            case '[': return ParseArray(depth, at);
            case '"': return JsonData.CreateString(ParseString(), at);
            case 't': ExpectWord("true"); return JsonData.CreateBoolean(true, at);
            case 'f': ExpectWord("false"); return JsonData.CreateBoolean(false, at);
            case 'n': ExpectWord("null"); return JsonData.CreateNull(at);
            default:
                if (c == '-' || (c >= '0' && c <= '9')) return ParseNumber(at);
                throw Error("Unexpected character '" + c + "'");
        }
    }

    private JsonData ParseObject(int depth, JsonDataLocation at) {
        _pos++;
        var properties = new List<KeyValuePair<string, JsonData>>();
        SkipWhitespace();
        if (Peek() == '}') {
            _pos++;
            return JsonData.CreateObject(properties, at);
        }
        while (true) {
            SkipWhitespace();
            if (Peek() != '"') throw Error("Expected a property name in double quotes");
            string name = ParseString();
            SkipWhitespace();
            if (Peek() != ':') throw Error("Expected ':' after a property name");
            _pos++;
            SkipWhitespace();
            properties.Add(new KeyValuePair<string, JsonData>(name, ParseValue(depth + 1)));
            SkipWhitespace();
            char next = Peek();
            _pos++;
            if (next == ',') continue;
            if (next == '}') return JsonData.CreateObject(properties, at);
            _pos--;
            throw Error("Expected ',' or '}' in an object");
        }
    }

    private JsonData ParseArray(int depth, JsonDataLocation at) {
        _pos++;
        var items = new List<JsonData>();
        SkipWhitespace();
        if (Peek() == ']') {
            _pos++;
            return JsonData.CreateArray(items, at);
        }
        while (true) {
            SkipWhitespace();
            items.Add(ParseValue(depth + 1));
            SkipWhitespace();
            char next = Peek();
            _pos++;
            if (next == ',') continue;
            if (next == ']') return JsonData.CreateArray(items, at);
            _pos--;
            throw Error("Expected ',' or ']' in an array");
        }
    }

    private string ParseString() {
        _pos++;
        int start = _pos;
        while (_pos < _text.Length) {
            char c = _text[_pos];
            if (c == '"') {
                string simple = _text.Substring(start, _pos - start);
                _pos++;
                return simple;
            }
            if (c == '\\') break;
            if (c < 0x20) throw Error("Control characters must be escaped in strings");
            _pos++;
        }
        var text = new StringBuilder(_text, start, _pos - start, Math.Max(16, (_pos - start) * 2));
        while (_pos < _text.Length) {
            char c = _text[_pos++];
            if (c == '"') return text.ToString();
            if (c < 0x20) {
                _pos--;
                throw Error("Control characters must be escaped in strings");
            }
            if (c != '\\') {
                text.Append(c);
                continue;
            }
            if (_pos >= _text.Length) break;
            char escape = _text[_pos++];
            switch (escape) {
                case '"': text.Append('"'); break;
                case '\\': text.Append('\\'); break;
                case '/': text.Append('/'); break;
                case 'b': text.Append('\b'); break;
                case 'f': text.Append('\f'); break;
                case 'n': text.Append('\n'); break;
                case 'r': text.Append('\r'); break;
                case 't': text.Append('\t'); break;
                case 'u':
                    if (_pos + 4 > _text.Length || !int.TryParse(_text.Substring(_pos, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int code)) {
                        throw Error("Invalid \\u escape");
                    }
                    text.Append((char)code);
                    _pos += 4;
                    break;
                default:
                    _pos--;
                    throw Error("Invalid escape '\\" + escape + "'");
            }
        }
        throw Error("Unterminated string");
    }

    private JsonData ParseNumber(JsonDataLocation at) {
        int start = _pos;
        if (Peek() == '-') _pos++;
        if (Peek() == '0') {
            _pos++;
        } else if (IsDigit(Peek())) {
            while (IsDigit(Peek())) _pos++;
        } else {
            throw Error("Invalid number");
        }
        if (Peek() == '.') {
            _pos++;
            if (!IsDigit(Peek())) throw Error("Expected a digit after the decimal point");
            while (IsDigit(Peek())) _pos++;
        }
        if (Peek() is 'e' or 'E') {
            _pos++;
            if (Peek() is '+' or '-') _pos++;
            if (!IsDigit(Peek())) throw Error("Expected a digit in the exponent");
            while (IsDigit(Peek())) _pos++;
        }
        return JsonData.ParseNumber(_text.Substring(start, _pos - start), at);
    }

    private void ExpectWord(string word) {
        if (string.CompareOrdinal(_text, _pos, word, 0, word.Length) != 0) throw Error("Unexpected token");
        _pos += word.Length;
    }

    private void SkipWhitespace() {
        while (_pos < _text.Length) {
            char c = _text[_pos];
            if (c == '\n') {
                _line++;
                _lineStart = _pos + 1;
            } else if (c == '\r') {
                if (_pos + 1 < _text.Length && _text[_pos + 1] == '\n') _pos++;
                _line++;
                _lineStart = _pos + 1;
            } else if (c != ' ' && c != '\t') {
                return;
            }
            _pos++;
        }
    }

    private char Peek() => _pos < _text.Length ? _text[_pos] : '\0';

    private static bool IsDigit(char c) => c >= '0' && c <= '9';

    private FormatException Error(string message) =>
        new FormatException(message + " at line " + _line.ToString(CultureInfo.InvariantCulture) + ", column " + (_pos - _lineStart + 1).ToString(CultureInfo.InvariantCulture) + ".");
}
