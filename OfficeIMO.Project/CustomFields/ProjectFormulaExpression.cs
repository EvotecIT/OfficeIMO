using System.Globalization;
using System.Text;

namespace OfficeIMO.Project;

/// <summary>Bounded, non-executable expression parser. It never calls VBA, reflection, COM, or external code.</summary>
internal sealed partial class ProjectFormulaExpression {
    private readonly string _source;
    private readonly Func<string, ProjectFormulaValue> _field;
    private readonly CancellationToken _token;
    private readonly int _maxOperations;
    private readonly ProjectFormulaNesting _nesting;
    private readonly CultureInfo _culture;
    private int _position, _operations;
    private string _kind = "", _text = "";
    internal ProjectFormulaExpression(string source, Func<string, ProjectFormulaValue> field, int maxOperations, ProjectFormulaNesting nesting, CultureInfo culture, CancellationToken token) {
        _source = source; _field = field; _maxOperations = maxOperations; _nesting = nesting; _token = token;
        _culture = culture;
        if (source.Length > 65536) throw new InvalidDataException("A formula exceeds the 65536-character limit.");
    }
    internal ProjectFormulaValue Evaluate() {
        Next(); var value = Expression(0);
        if (_kind != "end") throw Invalid("Unexpected token after the expression");
        return value;
    }
    private InvalidDataException Invalid(string message) => new InvalidDataException(message + " at formula character " + _position.ToString(CultureInfo.InvariantCulture) + ".");
    private void Step() {
        _token.ThrowIfCancellationRequested();
        if (++_operations > _maxOperations) throw Invalid("Formula operation limit exceeded");
    }
    private ProjectFormulaValue Expression(int minimum) {
        Step();
        _nesting.Enter();
        try {
            var left = Prefix();
            while (Precedence(_kind) >= minimum) {
                string operation = _kind; int precedence = Precedence(operation); Next();
                var right = Expression(precedence + 1);
                left = Binary(operation, left, right);
            }
            return left;
        } finally { _nesting.Leave(); }
    }
    private ProjectFormulaValue Prefix() {
        string kind = _kind, text = _text;
        Next();
        if (kind == "+" || kind == "-") {
            decimal number = Expression(9).NumberIn(_culture);
            return new ProjectFormulaValue(kind == "-" ? -number : number);
        }
        if (kind == "not") return new ProjectFormulaValue(!Expression(3).Logical);
        if (kind == "number") return new ProjectFormulaValue(decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
        if (kind == "string") return new ProjectFormulaValue(text);
        if (kind == "field") return _field(text);
        if (kind == "(") { var value = Expression(0); Require(")"); return value; }
        if (kind == "name") {
            if (text.Equals("True", StringComparison.OrdinalIgnoreCase)) return new ProjectFormulaValue(true);
            if (text.Equals("False", StringComparison.OrdinalIgnoreCase)) return new ProjectFormulaValue(false);
            Require("("); var arguments = new List<ProjectFormulaValue>();
            if (_kind != ")") {
                do {
                    if (arguments.Count == 32) throw Invalid("A function exceeds the argument limit");
                    arguments.Add(Expression(0));
                    if (_kind != ",") break;
                    Next();
                } while (true);
            }
            Require(")"); return Call(text, arguments);
        }
        throw Invalid("Expected a literal, field reference, function, or parenthesized expression");
    }
    private void Require(string kind) { if (_kind != kind) throw Invalid("Expected '" + kind + "'"); Next(); }
    private static int Precedence(string operation) => operation switch {
        "xor" => 0, "or" => 1, "and" => 2,
        "=" or "<>" or "<" or ">" or "<=" or ">=" => 3,
        "&" => 4, "+" or "-" => 5, "mod" => 6, "\\" => 7, "*" or "/" => 8, "^" => 10, _ => -1
    };
    private ProjectFormulaValue Binary(string operation, ProjectFormulaValue left, ProjectFormulaValue right) {
        return new ProjectFormulaValue(operation switch {
            "+" => left.NumberIn(_culture) + right.NumberIn(_culture), "-" => left.NumberIn(_culture) - right.NumberIn(_culture),
            "*" => left.NumberIn(_culture) * right.NumberIn(_culture), "/" => left.NumberIn(_culture) / right.NumberIn(_culture),
            "\\" => decimal.Truncate(decimal.Round(left.NumberIn(_culture), 0, MidpointRounding.ToEven) / decimal.Round(right.NumberIn(_culture), 0, MidpointRounding.ToEven)),
            "mod" => decimal.Round(left.NumberIn(_culture), 0, MidpointRounding.ToEven) % decimal.Round(right.NumberIn(_culture), 0, MidpointRounding.ToEven),
            "^" => (decimal)Math.Pow((double)left.NumberIn(_culture), (double)right.NumberIn(_culture)),
            "&" => Join(left.TextIn(_culture), right.TextIn(_culture)),
            "=" => ProjectFormulaValue.Compare(left, right, _culture) == 0, "<>" => ProjectFormulaValue.Compare(left, right, _culture) != 0,
            "<" => ProjectFormulaValue.Compare(left, right, _culture) < 0, ">" => ProjectFormulaValue.Compare(left, right, _culture) > 0,
            "<=" => ProjectFormulaValue.Compare(left, right, _culture) <= 0, ">=" => ProjectFormulaValue.Compare(left, right, _culture) >= 0,
            "and" => left.Logical & right.Logical, "or" => left.Logical | right.Logical, "xor" => left.Logical ^ right.Logical,
            _ => throw new NotSupportedException("The expression operator is not supported.")
        });
    }
    private static string Join(string left, string right) {
        if ((long)left.Length + right.Length > 65536) throw new InvalidDataException("Formula text output exceeds 65536 characters.");
        return left + right;
    }
    private void Next() {
        Step();
        while (_position < _source.Length && char.IsWhiteSpace(_source[_position])) _position++;
        if (_position == _source.Length) { _kind = "end"; _text = ""; return; }
        char c = _source[_position++]; int start = _position - 1;
        if (c == '[') {
            int close = _source.IndexOf(']', _position);
            if (close < 0) throw Invalid("Unclosed field reference");
            _text = _source.Substring(_position, close - _position); _position = close + 1; _kind = "field"; return;
        }
        if (c == '"') {
            var text = new StringBuilder();
            while (_position < _source.Length) {
                char next = _source[_position++];
                if (next != '"') { text.Append(next); continue; }
                if (_position < _source.Length && _source[_position] == '"') { text.Append('"'); _position++; continue; }
                _kind = "string"; _text = text.ToString(); return;
            }
            throw Invalid("Unclosed string literal");
        }
        if (char.IsDigit(c) || c == '.') {
            bool dot = c == '.', exponent = false;
            while (_position < _source.Length) {
                char next = _source[_position];
                if (char.IsDigit(next)) { _position++; continue; }
                if (next == '.' && !dot && !exponent) { dot = true; _position++; continue; }
                if ((next == 'e' || next == 'E') && !exponent) {
                    exponent = true; _position++;
                    if (_position < _source.Length && (_source[_position] == '+' || _source[_position] == '-')) _position++;
                    continue;
                }
                break;
            }
            _text = _source.Substring(start, _position - start); _kind = "number"; return;
        }
        if (char.IsLetter(c) || c == '_') {
            while (_position < _source.Length && (char.IsLetterOrDigit(_source[_position]) || _source[_position] == '_')) _position++;
            _text = _source.Substring(start, _position - start); string lower = _text.ToLowerInvariant();
            _kind = lower is "and" or "or" or "xor" or "not" or "mod" ? lower : "name"; return;
        }
        _kind = c.ToString(); _text = _kind;
        if ((c == '<' || c == '>') && _position < _source.Length && (_source[_position] == '=' || c == '<' && _source[_position] == '>'))
            _kind += _source[_position++];
        if ("+-*/\\^&=<>(),".IndexOf(c) < 0) throw Invalid("Unsupported formula character");
    }
}

// Shared across dependent formulas so individually shallow expressions cannot multiply the permitted stack depth.
internal sealed class ProjectFormulaNesting {
    private readonly int _maximum;
    private int _depth;
    internal ProjectFormulaNesting(int maximum) { _maximum = maximum; }
    internal void Enter() {
        if (_depth >= _maximum) throw new InvalidDataException("Formula nesting limit exceeded.");
        _depth++;
    }
    internal void Leave() { _depth--; }
}
