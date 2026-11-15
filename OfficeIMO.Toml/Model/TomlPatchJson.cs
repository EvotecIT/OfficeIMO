using System.Collections.Specialized;

namespace OfficeIMO.Toml;

/// <summary>
/// Reads the JSON text of a patch into plain values: <see cref="OrderedDictionary"/> for objects (member order kept),
/// <see cref="List{T}"/> for arrays, <see cref="string"/>, <see cref="long"/> (or <see cref="double"/> for fractions,
/// exponents and integers beyond 64 bits), <see cref="bool"/> and <see langword="null"/>. The engine has no JSON
/// dependency, and a patch needs no more than this.
/// </summary>
internal sealed class TomlPatchJson {
    private const int MaxDepth = 256;

    private readonly string _text;
    private int _position;

    private TomlPatchJson(string text) => _text = text;

    /// <exception cref="TomlPatchException">The text is not one JSON value, or an object repeats a member name.</exception>
    public static object? Parse(string text) {
        var reader = new TomlPatchJson(text);
        reader.SkipSpace();
        object? value = reader.ReadValue(0);
        reader.SkipSpace();
        if (reader._position != text.Length) throw reader.Error("unexpected text after the value");
        return value;
    }

    private object? ReadValue(int depth) {
        if (depth > MaxDepth) throw Error("the value nests more than " + MaxDepth + " levels deep");
        if (_position >= _text.Length) throw Error("a value is missing");
        char c = _text[_position];
        switch (c) {
            case '{': return ReadObject(depth);
            case '[': return ReadArray(depth);
            case '"': return ReadString();
            case 't': Expect("true"); return true;
            case 'f': Expect("false"); return false;
            case 'n': Expect("null"); return null;
            default:
                if (c == '-' || (c >= '0' && c <= '9')) return ReadNumber();
                throw Error("'" + c + "' does not start a value");
        }
    }

    private OrderedDictionary ReadObject(int depth) {
        var members = new OrderedDictionary(StringComparer.Ordinal);
        _position++;
        SkipSpace();
        if (Peek() == '}') {
            _position++;
            return members;
        }
        while (true) {
            SkipSpace();
            if (Peek() != '"') throw Error("a member name is a string");
            int start = _position;
            string name = ReadString();
            if (members.Contains(name)) {
                _position = start;
                throw Error("the member '" + name + "' appears twice");
            }
            SkipSpace();
            if (Peek() != ':') throw Error("':' follows a member name");
            _position++;
            SkipSpace();
            members.Add(name, ReadValue(depth + 1));
            SkipSpace();
            char next = Peek();
            _position++;
            if (next == '}') return members;
            if (next != ',') {
                _position--;
                throw Error("',' or '}' follows a member");
            }
        }
    }

    private List<object?> ReadArray(int depth) {
        var items = new List<object?>();
        _position++;
        SkipSpace();
        if (Peek() == ']') {
            _position++;
            return items;
        }
        while (true) {
            SkipSpace();
            items.Add(ReadValue(depth + 1));
            SkipSpace();
            char next = Peek();
            _position++;
            if (next == ']') return items;
            if (next != ',') {
                _position--;
                throw Error("',' or ']' follows an item");
            }
        }
    }

    private string ReadString() {
        _position++;
        var text = new StringBuilder();
        while (true) {
            if (_position >= _text.Length) throw Error("the string is not closed");
            char c = _text[_position++];
            if (c == '"') return text.ToString();
            if (c < ' ') {
                _position--;
                throw Error("a control character in a string is escaped");
            }
            if (c != '\\') {
                text.Append(c);
                continue;
            }
            if (_position >= _text.Length) throw Error("the string is not closed");
            char escape = _text[_position++];
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
                    if (_position + 4 > _text.Length
                        || !int.TryParse(_text.Substring(_position, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int code)) {
                        throw Error("'\\u' is followed by four hexadecimal digits");
                    }
                    text.Append((char)code);
                    _position += 4;
                    break;
                default:
                    _position--;
                    throw Error("'\\" + escape + "' is not a JSON escape");
            }
        }
    }

    private object ReadNumber() {
        int start = _position;
        bool integer = true;
        if (Peek() == '-') _position++;
        if (Peek() == '0') _position++;
        else if (!Digits()) throw Error("a number has digits");
        if (Peek() == '.') {
            integer = false;
            _position++;
            if (!Digits()) throw Error("a fraction has digits");
        }
        if (Peek() is 'e' or 'E') {
            integer = false;
            _position++;
            if (Peek() is '+' or '-') _position++;
            if (!Digits()) throw Error("an exponent has digits");
        }
        string text = _text.Substring(start, _position - start);
        if (integer && long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long whole)) return whole;
        double number;
        try {
            number = double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture);
        } catch (OverflowException) {
            number = double.PositiveInfinity;
        }
        // .NET Framework throws for a number beyond double's range and later runtimes return infinity; both are rejected.
        if (double.IsInfinity(number)) {
            _position = start;
            throw Error("the number is beyond the range of a double");
        }
        return number;
    }

    private bool Digits() {
        int start = _position;
        while (_position < _text.Length && _text[_position] >= '0' && _text[_position] <= '9') _position++;
        return _position > start;
    }

    private void Expect(string literal) {
        if (string.CompareOrdinal(_text, _position, literal, 0, literal.Length) != 0) throw Error("'" + literal + "' is misspelled");
        _position += literal.Length;
    }

    private char Peek() => _position < _text.Length ? _text[_position] : '\0';

    private void SkipSpace() {
        while (_position < _text.Length && _text[_position] is ' ' or '\t' or '\n' or '\r') _position++;
    }

    private TomlPatchException Error(string reason) =>
        new TomlPatchException(-1, null, "The patch is not valid JSON at offset " + _position.ToString(CultureInfo.InvariantCulture) + ": " + reason + ".", null);
}
