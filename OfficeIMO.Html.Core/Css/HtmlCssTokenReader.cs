using System;
using System.Text;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Iterative CSS Syntax lexical reader; all offsets refer to the unmodified input.</summary>
internal sealed class HtmlCssTokenReader {
    private readonly string _source;
    private readonly CancellationToken _cancellation;
    private int _position;

    internal HtmlCssTokenReader(string source, CancellationToken cancellation) {
        _source = source;
        _cancellation = cancellation;
    }

    internal HtmlCssToken Read() {
        _cancellation.ThrowIfCancellationRequested();
        int start = _position;
        char current = Peek();
        if (AtEnd) return Token(HtmlCssTokenKind.EndOfFile, start);
        if (Whitespace(current)) {
            do { Advance(); } while (Whitespace(Peek()));
            return Token(HtmlCssTokenKind.Whitespace, start);
        }
        if (current == '/' && Peek(1) == '*') {
            Advance(); Advance();
            while (!AtEnd && !(Peek() == '*' && Peek(1) == '/')) Advance();
            if (!AtEnd) { Advance(); Advance(); }
            return Token(HtmlCssTokenKind.Comment, start);
        }
        if (current == '"' || current == '\'') return ReadString(start);
        if (StartsNumber(_position)) return ReadNumeric(start);
        if (current == '-' && Peek(1) == '-' && Peek(2) == '>') {
            Advance(); Advance(); Advance();
            return Token(HtmlCssTokenKind.Cdc, start);
        }
        if (StartsIdentifier(_position)) {
            string name = ReadName();
            if (Peek() != '(') return Token(HtmlCssTokenKind.Identifier, start, name);
            Advance();
            if (string.Equals(name, "url", StringComparison.OrdinalIgnoreCase)) {
                int lookahead = _position;
                while (Whitespace(At(lookahead))) {
                    _cancellation.ThrowIfCancellationRequested();
                    lookahead++;
                }
                if (At(lookahead) != '"' && At(lookahead) != '\'') return ReadUrl(start);
            }
            return Token(HtmlCssTokenKind.Function, start, name);
        }
        if (current == '#' && (Name(Peek(1)) || ValidEscape(_position + 1))) {
            Advance();
            bool identifier = StartsIdentifier(_position);
            return Token(HtmlCssTokenKind.Hash, start, ReadName(), identifier);
        }
        if (current == '@' && StartsIdentifier(_position + 1)) {
            Advance();
            return Token(HtmlCssTokenKind.AtKeyword, start, ReadName());
        }
        if (current == '<' && Peek(1) == '!' && Peek(2) == '-' && Peek(3) == '-') {
            for (int index = 0; index < 4; index++) Advance();
            return Token(HtmlCssTokenKind.Cdo, start);
        }
        Advance();
        HtmlCssTokenKind kind;
        switch (current) {
            case ':': kind = HtmlCssTokenKind.Colon; break;
            case ';': kind = HtmlCssTokenKind.Semicolon; break;
            case ',': kind = HtmlCssTokenKind.Comma; break;
            case '(': kind = HtmlCssTokenKind.OpenParenthesis; break;
            case ')': kind = HtmlCssTokenKind.CloseParenthesis; break;
            case '[': kind = HtmlCssTokenKind.OpenBracket; break;
            case ']': kind = HtmlCssTokenKind.CloseBracket; break;
            case '{': kind = HtmlCssTokenKind.OpenBrace; break;
            case '}': kind = HtmlCssTokenKind.CloseBrace; break;
            default: kind = HtmlCssTokenKind.Delimiter; break;
        }
        return Token(kind, start, kind == HtmlCssTokenKind.Delimiter ? current.ToString() : null);
    }

    private HtmlCssToken ReadString(int start) {
        char quote = Peek();
        Advance();
        var value = new StringBuilder();
        while (!AtEnd) {
            char current = Peek();
            if (current == quote) { Advance(); return Token(HtmlCssTokenKind.String, start, value.ToString()); }
            if (current == '\n') return Token(HtmlCssTokenKind.BadString, start);
            if (current == '\\') {
                Advance();
                if (AtEnd) break;
                if (Peek() == '\n') Advance();
                else value.Append(ReadEscape());
            } else {
                value.Append(current);
                Advance();
            }
        }
        return Token(HtmlCssTokenKind.String, start, value.ToString());
    }

    private HtmlCssToken ReadUrl(int start) {
        while (Whitespace(Peek())) Advance();
        var value = new StringBuilder();
        while (!AtEnd) {
            char current = Peek();
            if (current == ')') { Advance(); return Token(HtmlCssTokenKind.Url, start, value.ToString()); }
            if (Whitespace(current)) {
                do { Advance(); } while (Whitespace(Peek()));
                if (AtEnd) break;
                if (Peek() == ')') { Advance(); return Token(HtmlCssTokenKind.Url, start, value.ToString()); }
                return ReadBadUrl(start);
            }
            if (current == '"' || current == '\'' || current == '(' || NonPrintable(current)) return ReadBadUrl(start);
            if (current == '\\') {
                if (!ValidEscape(_position)) return ReadBadUrl(start);
                Advance();
                value.Append(ReadEscape());
            } else {
                value.Append(current);
                Advance();
            }
        }
        return Token(HtmlCssTokenKind.Url, start, value.ToString());
    }

    private HtmlCssToken ReadBadUrl(int start) {
        while (!AtEnd) {
            if (Peek() == ')') { Advance(); break; }
            if (ValidEscape(_position)) { Advance(); ReadEscape(); }
            else Advance();
        }
        return Token(HtmlCssTokenKind.BadUrl, start);
    }

    private HtmlCssToken ReadNumeric(int start) {
        if (Peek() == '+' || Peek() == '-') Advance();
        while (Digit(Peek())) Advance();
        if (Peek() == '.' && Digit(Peek(1))) {
            Advance();
            while (Digit(Peek())) Advance();
        }
        if ((Peek() == 'e' || Peek() == 'E') && (Digit(Peek(1)) || ((Peek(1) == '+' || Peek(1) == '-') && Digit(Peek(2))))) {
            Advance();
            if (Peek() == '+' || Peek() == '-') Advance();
            while (Digit(Peek())) Advance();
        }
        if (StartsIdentifier(_position)) return Token(HtmlCssTokenKind.Dimension, start, ReadName());
        if (Peek() == '%') { Advance(); return Token(HtmlCssTokenKind.Percentage, start); }
        return Token(HtmlCssTokenKind.Number, start);
    }

    private string ReadName() {
        var value = new StringBuilder();
        while (!AtEnd) {
            if (Name(Peek())) { value.Append(Peek()); Advance(); }
            else if (ValidEscape(_position)) { Advance(); value.Append(ReadEscape()); }
            else break;
        }
        return value.ToString();
    }

    // Called after consuming the backslash. A trailing backslash becomes U+FFFD.
    private string ReadEscape() {
        if (AtEnd) return "\ufffd";
        if (!Hex(Peek())) {
            char current = Peek();
            // Preserve a valid non-BMP code point escaped as a literal surrogate pair.
            if (char.IsHighSurrogate(current) && char.IsLowSurrogate(Peek(1))) {
                string pair = _source.Substring(_position, 2);
                Advance(); Advance();
                return pair;
            }
            Advance();
            return current.ToString();
        }
        int scalar = 0;
        for (int count = 0; count < 6 && Hex(Peek()); count++) {
            char current = Peek();
            scalar = scalar * 16 + (current <= '9' ? current - '0' : char.ToLowerInvariant(current) - 'a' + 10);
            Advance();
        }
        if (Whitespace(Peek())) Advance();
        return scalar == 0 || scalar > 0x10ffff || scalar >= 0xd800 && scalar <= 0xdfff
            ? "\ufffd" : char.ConvertFromUtf32(scalar);
    }

    private bool StartsIdentifier(int position) {
        char first = At(position);
        if (first == '-') return NameStart(At(position + 1)) || At(position + 1) == '-' || ValidEscape(position + 1);
        return NameStart(first) || ValidEscape(position);
    }

    private bool StartsNumber(int position) {
        char first = At(position);
        if (first == '+' || first == '-') return Digit(At(position + 1)) || At(position + 1) == '.' && Digit(At(position + 2));
        return Digit(first) || first == '.' && Digit(At(position + 1));
    }

    private bool ValidEscape(int position) => At(position) == '\\' && At(position + 1) != '\n';
    private bool AtEnd => _position >= _source.Length;
    private char Peek(int ahead = 0) => At(_position + ahead);
    private char At(int position) {
        if (position >= _source.Length) return '\0';
        char current = _source[position];
        if (current == '\r' || current == '\f') return '\n';
        if (current == '\0') return '\ufffd';
        if (char.IsHighSurrogate(current) && (position + 1 == _source.Length || !char.IsLowSurrogate(_source[position + 1]))) return '\ufffd';
        if (char.IsLowSurrogate(current) && (position == 0 || !char.IsHighSurrogate(_source[position - 1]))) return '\ufffd';
        return current;
    }

    private void Advance() {
        _cancellation.ThrowIfCancellationRequested();
        if (AtEnd) return;
        if (_source[_position++] == '\r' && !AtEnd && _source[_position] == '\n') _position++;
    }

    private HtmlCssToken Token(HtmlCssTokenKind kind, int start, string? value = null, bool identifierHash = false) =>
        new HtmlCssToken(kind, start, _position - start, value, identifierHash);
    private static bool Whitespace(char value) => value == ' ' || value == '\t' || value == '\n';
    private static bool Digit(char value) => value >= '0' && value <= '9';
    private static bool Hex(char value) => Digit(value) || value >= 'a' && value <= 'f' || value >= 'A' && value <= 'F';
    private static bool NameStart(char value) => value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' || value == '_' || value >= 0x80;
    private static bool Name(char value) => NameStart(value) || Digit(value) || value == '-';
    private static bool NonPrintable(char value) => value >= '\u0001' && value <= '\u0008' || value == '\u000b' || value >= '\u000e' && value <= '\u001f' || value == '\u007f';
}
