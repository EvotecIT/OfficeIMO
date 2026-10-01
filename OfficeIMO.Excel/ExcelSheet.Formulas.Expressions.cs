using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateScalarExpression(string expression, out FormulaArgumentValue value) {
            var parser = new ScalarFormulaParser(this, expression);
            return parser.TryParse(out value);
        }

        // Single-pass precedence parsing. References and functions use the existing evaluator;
        // syntax nesting and function recursion have separate bounded guards.
        private struct ScalarFormulaParser {
            private readonly ExcelSheet _sheet;
            private readonly string _text;
            private int _position;
            private int _depth;

            internal ScalarFormulaParser(ExcelSheet sheet, string text) {
                _sheet = sheet;
                _text = text;
                _position = 0;
                _depth = 0;
            }

            internal bool TryParse(out FormulaArgumentValue value) {
                if (!Parse(1, out value)) return false;
                SkipWhiteSpace();
                return _position == _text.Length && !value.IsUnresolvedFormula;
            }

            private bool Parse(int minimumPrecedence, out FormulaArgumentValue value) {
                value = default;
                if (!_sheet.HasSufficientFormulaExecutionStack()) return false;
                if (++_depth > 128) { _depth--; return false; }
                try {
                    SkipWhiteSpace();
                    if (_position == _text.Length) return false;
                    char first = _text[_position];
                    if (first == '+' || first == '-') {
                        _position++;
                        if (!Parse(7, out value)) return false;
                        value = NumericUnary(value, first == '-' ? -1 : 1);
                    } else if (first == '(') {
                        _position++;
                        if (!Parse(1, out value)) return false;
                        SkipWhiteSpace();
                        if (_position == _text.Length || _text[_position++] != ')') return false;
                    } else if (!ReadAtom(out value)) {
                        return false;
                    }

                    while (true) {
                        SkipWhiteSpace();
                        if (_position == _text.Length) break;
                        string op = ReadOperator();
                        int precedence = Precedence(op);
                        if (precedence < minimumPrecedence) break;
                        _position += op.Length;
                        if (op == "%") {
                            value = NumericUnary(value, 0.01);
                            continue;
                        }
                        // Excel evaluates equal-precedence operators left to right, including ^.
                        if (!Parse(precedence + 1, out FormulaArgumentValue right)) return false;
                        if (!Apply(value, op, right, out value)) return false;
                    }
                    return !value.IsUnresolvedFormula;
                } finally { _depth--; }
            }

            private bool ReadAtom(out FormulaArgumentValue value) {
                value = default;
                int start = _position;
                if (_text[start] == '#') {
                    foreach (string error in ScalarErrorLiterals) {
                        if (_text.Length - start >= error.Length && string.Compare(_text, start, error, 0, error.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                            _position += error.Length;
                            value = FormulaArgumentValue.Error(error);
                            return true;
                        }
                    }
                }
                int parentheses = 0, brackets = 0;
                char quote = '\0';
                while (_position < _text.Length) {
                    char ch = _text[_position];
                    if (quote == '\0' && brackets > 0 && ch == '\'' && _position + 1 < _text.Length) {
                        _position += 2;
                        continue;
                    }
                    if (quote != '\0') {
                        if (ch == quote) {
                            if (_position + 1 < _text.Length && _text[_position + 1] == quote) {
                                _position += 2;
                                continue;
                            }
                            quote = '\0';
                        }
                    } else if (brackets == 0 && (ch == '"' || ch == '\'')) {
                        quote = ch;
                    } else if (ch == '[') { brackets++; }
                    else if (ch == ']') { if (--brackets < 0) return false; }
                    else if (brackets == 0) {
                        if (ch == '(') { if (++parentheses > 128) return false; }
                        else if (ch == ')') { if (parentheses == 0) break; parentheses--; }
                        else if (parentheses == 0 && IsOperator(ch)) {
                            // A sign in a scientific numeric literal belongs to the literal.
                            bool exponentSign = (ch == '+' || ch == '-') && _position > start
                                && (_text[_position - 1] == 'e' || _text[_position - 1] == 'E')
                                && double.TryParse(_text.Substring(start, _position - start - 1), NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out _);
                            if (!exponentSign) break;
                        }
                    }
                    _position++;
                }
                if (quote != '\0' || parentheses != 0 || brackets != 0 || start == _position) return false;
                return _sheet.TryResolveFormulaArgument(_text.Substring(start, _position - start), out value,
                    allowScalarExpression: false);
            }

            private static bool IsOperator(char ch) => ch == '+' || ch == '-' || ch == '*' || ch == '/'
                || ch == '^' || ch == '%' || ch == '&' || ch == '=' || ch == '<' || ch == '>';

            private static readonly string[] ScalarErrorLiterals = { "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!" };

            private string ReadOperator() {
                char ch = _text[_position];
                if ((ch == '<' || ch == '>') && _position + 1 < _text.Length) {
                    char next = _text[_position + 1];
                    if (next == '=') return ch == '<' ? "<=" : ">=";
                    if (ch == '<' && next == '>') return "<>";
                }
                switch (ch) {
                    case '+': return "+";
                    case '-': return "-";
                    case '*': return "*";
                    case '/': return "/";
                    case '^': return "^";
                    case '%': return "%";
                    case '&': return "&";
                    case '=': return "=";
                    case '<': return "<";
                    case '>': return ">";
                    default: return string.Empty;
                }
            }

            private static int Precedence(string op) {
                switch (op) {
                    case "=": case "<>": case "<": case ">": case "<=": case ">=": return 1;
                    case "&": return 2;
                    case "+": case "-": return 3;
                    case "*": case "/": return 4;
                    case "^": return 5;
                    case "%": return 6;
                    default: return 0;
                }
            }

            private static bool TryNumber(FormulaArgumentValue value, out double number) {
                number = value.Number ?? 0;
                if (value.Number.HasValue || !value.HasValue) return true;
                return double.TryParse(value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
            }

            private static FormulaArgumentValue Numeric(double number) =>
                double.IsNaN(number) || double.IsInfinity(number) ? FormulaArgumentValue.Error("#NUM!")
                    : new FormulaArgumentValue(number, InvariantNumberText.Get(number));

            private static FormulaArgumentValue NumericUnary(FormulaArgumentValue value, double factor) {
                if (value.IsError || value.IsUnresolvedFormula) return value;
                return TryNumber(value, out double number) ? Numeric(number * factor) : FormulaArgumentValue.Error("#VALUE!");
            }

            private static bool Apply(FormulaArgumentValue left, string op, FormulaArgumentValue right,
                out FormulaArgumentValue value) {
                value = default;
                if (left.IsUnresolvedFormula || right.IsUnresolvedFormula) return false;
                if (left.IsError || right.IsError) { value = left.IsError ? left : right; return true; }
                if (op == "&") { value = new FormulaArgumentValue(null, ScalarText(left) + ScalarText(right)); return true; }
                if (Precedence(op) == 1) {
                    if (!TryCompareFormulaValues(left, op, right, out bool comparison)) return false;
                    value = new FormulaArgumentValue(comparison ? 1 : 0, comparison ? "1" : "0", isBoolean: true);
                    return true;
                }
                if (!TryNumber(left, out double a) || !TryNumber(right, out double b)) {
                    value = FormulaArgumentValue.Error("#VALUE!"); return true;
                }
                switch (op) {
                    case "+": value = Numeric(a + b); break;
                    case "-": value = Numeric(a - b); break;
                    case "*": value = Numeric(a * b); break;
                    case "/": value = b == 0 ? FormulaArgumentValue.Error("#DIV/0!") : Numeric(a / b); break;
                    case "^": value = Numeric(Math.Pow(a, b)); break;
                    default: return false;
                }
                return true;
            }

            private void SkipWhiteSpace() {
                while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++;
            }

            private static string ScalarText(FormulaArgumentValue value) => value.IsBoolean
                ? (value.Number == 0 ? "FALSE" : "TRUE")
                : value.Number.HasValue ? InvariantNumberText.Get(value.Number.Value) : FormulaValueToText(value);
        }
    }
}
