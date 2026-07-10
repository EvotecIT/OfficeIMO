namespace OfficeIMO.OpenDocument;

internal enum OdsFormulaTokenKind {
    End, Number, String, Identifier, Reference,
    Plus, Minus, Star, Slash, Caret, Ampersand, Percent,
    Equal, NotEqual, Less, LessOrEqual, Greater, GreaterOrEqual,
    LeftParenthesis, RightParenthesis, Separator
}

internal readonly struct OdsFormulaToken {
    internal OdsFormulaToken(OdsFormulaTokenKind kind, string text) { Kind = kind; Text = text; }
    internal OdsFormulaTokenKind Kind { get; }
    internal string Text { get; }
}

internal sealed class OdsFormulaLexer {
    private readonly string _text;
    private int _position;
    internal OdsFormulaLexer(string text) { _text = text ?? string.Empty; }

    internal OdsFormulaToken Next() {
        while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++;
        if (_position >= _text.Length) return new OdsFormulaToken(OdsFormulaTokenKind.End, string.Empty);
        char current = _text[_position++];
        switch (current) {
            case '+': return Token(OdsFormulaTokenKind.Plus, current);
            case '-': return Token(OdsFormulaTokenKind.Minus, current);
            case '*': return Token(OdsFormulaTokenKind.Star, current);
            case '/': return Token(OdsFormulaTokenKind.Slash, current);
            case '^': return Token(OdsFormulaTokenKind.Caret, current);
            case '&': return Token(OdsFormulaTokenKind.Ampersand, current);
            case '%': return Token(OdsFormulaTokenKind.Percent, current);
            case '=': return Token(OdsFormulaTokenKind.Equal, current);
            case '(': return Token(OdsFormulaTokenKind.LeftParenthesis, current);
            case ')': return Token(OdsFormulaTokenKind.RightParenthesis, current);
            case ';':
            case ',': return Token(OdsFormulaTokenKind.Separator, current);
            case '<':
                if (TakeIf('=')) return new OdsFormulaToken(OdsFormulaTokenKind.LessOrEqual, "<=");
                if (TakeIf('>')) return new OdsFormulaToken(OdsFormulaTokenKind.NotEqual, "<>");
                return Token(OdsFormulaTokenKind.Less, current);
            case '>':
                return TakeIf('=') ? new OdsFormulaToken(OdsFormulaTokenKind.GreaterOrEqual, ">=") : Token(OdsFormulaTokenKind.Greater, current);
            case '"': return ReadString();
            case '[': return ReadReference();
            default:
                if (char.IsDigit(current) || current == '.') return ReadNumber(current);
                if (char.IsLetter(current) || current == '_') return ReadIdentifier(current);
                throw new OdsFormulaException("Unsupported formula character '" + current + "'.");
        }
    }

    private OdsFormulaToken ReadString() {
        var builder = new StringBuilder();
        while (_position < _text.Length) {
            char character = _text[_position++];
            if (character != '"') { builder.Append(character); continue; }
            if (_position < _text.Length && _text[_position] == '"') { builder.Append('"'); _position++; continue; }
            return new OdsFormulaToken(OdsFormulaTokenKind.String, builder.ToString());
        }
        throw new OdsFormulaException("Unterminated formula string.");
    }

    private OdsFormulaToken ReadReference() {
        int start = _position;
        bool quoted = false;
        while (_position < _text.Length) {
            char character = _text[_position++];
            if (character == '\'') {
                if (quoted && _position < _text.Length && _text[_position] == '\'') { _position++; continue; }
                quoted = !quoted;
            } else if (character == ']' && !quoted) {
                return new OdsFormulaToken(OdsFormulaTokenKind.Reference, _text.Substring(start, _position - start - 1));
            }
        }
        throw new OdsFormulaException("Unterminated formula reference.");
    }

    private OdsFormulaToken ReadNumber(char first) {
        int start = _position - 1;
        bool exponent = false;
        while (_position < _text.Length) {
            char character = _text[_position];
            if (char.IsDigit(character) || character == '.') { _position++; continue; }
            if ((character == 'e' || character == 'E') && !exponent) {
                exponent = true; _position++;
                if (_position < _text.Length && (_text[_position] == '+' || _text[_position] == '-')) _position++;
                continue;
            }
            break;
        }
        return new OdsFormulaToken(OdsFormulaTokenKind.Number, _text.Substring(start, _position - start));
    }

    private OdsFormulaToken ReadIdentifier(char first) {
        int start = _position - 1;
        while (_position < _text.Length && (char.IsLetterOrDigit(_text[_position]) || _text[_position] == '_' || _text[_position] == '.')) _position++;
        return new OdsFormulaToken(OdsFormulaTokenKind.Identifier, _text.Substring(start, _position - start));
    }

    private bool TakeIf(char value) {
        if (_position >= _text.Length || _text[_position] != value) return false;
        _position++; return true;
    }

    private static OdsFormulaToken Token(OdsFormulaTokenKind kind, char value) => new OdsFormulaToken(kind, value.ToString());
}
